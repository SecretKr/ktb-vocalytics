import argparse
import glob
import json
import locale
import logging
import os
import re
import sys
import time
from typing import Any, Dict, List, Optional, Tuple

from tqdm import tqdm
from tqdm.contrib.logging import _TqdmLoggingHandler as TqdmLoggingHandler
from tqdm.contrib.logging import logging_redirect_tqdm

# Set encoding
try:
    locale.setlocale(locale.LC_ALL, 'th_TH.utf-8')
    logging.info("Successfully set locale to th_TH.utf-8")
except locale.Error as e:
    logging.warning(f"Could not set locale to th_TH.utf-8: {e}")
except Exception as e:
    logging.warning(f"Unexpected error setting locale: {e}")


# ทำให้ log แสดงสวยร่วมกับ tqdm (ขึ้นเหนือ progress bar)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[TqdmLoggingHandler()],  # << สำคัญ
    force=True
)

# กดเสียง lib ภายนอก (ไม่ให้ HTTP 200 แทรกกลาง bar)
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("openai").setLevel(logging.WARNING)

import docx
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt, RGBColor
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv() # Ensure environment variables are loaded at the start

# Status constants for scoring
PASS_STATUSES = {"พบ"}

SCORING_LOGIC_NOTE = (
    "AI จะทำการวิเคราะห์และให้คะแนนตามเกณฑ์ที่กำหนด โดยคะแนนที่ได้จะถูกคำนวณจากเกณฑ์ทั้งหมด เพื่อให้ผลลัพธ์สะท้อนการทำงานได้อย่างครบถ้วน"
)

DECISION_GUIDE_TEXT = [
    "ดีมาก: คะแนนรวมสูง (เช่น ≥90%) อ้างอิงจากขั้นตอนที่พนักงานต้องปฏิบัติตามกระบวนการขาย",
    "ปานกลาง: คะแนนรวมอยู่ในระดับพอใช้ (เช่น 60–89%)  อ้างอิงจากขั้นตอนที่พนักงานต้องปฏิบัติตามกระบวนการขาย",
    "ควรปรับปรุง: คะแนนรวมต่ำ (<60%)  อ้างอิงจากขั้นตอนที่พนักงานต้องปฏิบัติตามกระบวนการขาย",
]

# --- DOCX File Handling ---

def _repeat_header_on_each_page(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader')
    trPr.append(tblHeader)


def read_docx(file_path: str) -> str:
    """Reads all text content from a .docx file, including paragraphs and tables."""
    try:
        doc = docx.Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    full_text.append(cell.text)
        return "\n".join(filter(None, full_text))
    except Exception as e:
        logging.error(f"Error reading DOCX file {file_path}: {e}")
        return ""

# --- Criteria Loading ---

def load_criteria_from_json(product_name: str, base_dir: str) -> Dict[str, Any]:
    """Loads product-specific criteria from a JSON file."""
    script_dir = os.path.dirname(__file__)
    path = os.path.join(script_dir, base_dir, f"{product_name}.json")
    if not os.path.exists(path):
        raise FileNotFoundError(f"Criteria file not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

# --- AI Interaction & Prompting ---

def get_openai_client() -> OpenAI:
    """Initializes and returns the OpenAI client for OpenRouter."""
    api_key = os.getenv("OPENROUTER_API_KEY")
    if not api_key:
        raise ValueError("OPENROUTER_API_KEY environment variable not set.")
    return OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=api_key,
        timeout=30.0,
    )

def is_rate_limit_error(e: Exception) -> bool:
    """Checks if an exception is a rate limit error (HTTP 429)."""
    return "429" in str(e)

def extract_json_from_response(text: str) -> Optional[Dict]:
    """
    ดึง JSON ที่ 'ใช่' จากคำตอบโมเดล:
    - ตัด code fences ออก
    - ลอง parse ตรง ๆ
    - ถ้าไม่สำเร็จ: แตกทุก {...} เป็น candidate แล้ว 'เลือกก้อนที่ดีที่สุด'
      โดยให้คะแนนตามโครงสร้างที่เราต้องการ (steps เป็น list ไม่ว่าง + มี summary)
    """
    if not text:
        return None

    s = text.strip()
    # ตัด ```...``` ออกถ้ามี
    s = re.sub(r"^```[a-zA-Z_]*\s*", "", s).strip()
    s = re.sub(r"\s*```$", "", s).strip()

    # 1) ลอง parse ตรง ๆ ก่อน
    try:
        obj = json.loads(s)
        return obj
    except Exception:
        pass

    # 2) แตกทุก {...} เป็น candidates
    candidates = [m.group(0) for m in re.finditer(r"\{.*?\}", s, flags=re.DOTALL)]
    best = None
    best_score = -1

    for cand in candidates:
        try:
            obj = json.loads(cand)
        except Exception:
            continue

        # ให้คะแนนความเหมาะสม
        score = 0
        steps = obj.get("steps")
        summary = obj.get("summary")
        if isinstance(steps, list) and len(steps) > 0:
            score += 2        # steps เป็น list และไม่ว่าง
            # เช็คว่า items ข้างในไม่ว่าง
            # (ไม่บังคับเสมอไป แต่ให้แต้มเพิ่ม)
            if any(isinstance(st.get("items"), list) and len(st.get("items")) > 0 for st in steps if isinstance(st, dict)):
                score += 1
        if isinstance(summary, dict):
            score += 1

        if score > best_score:
            best_score = score
            best = obj

    return best

def call_ai_model(client: OpenAI, messages: List[Dict[str, Any]], model: str, max_retries: int = 3) -> Optional[Dict[str, Any]]:
    """
    เรียกโมเดลแบบเรียบง่าย + backoff 429
    - ไม่พิมพ์เนื้อหาคำตอบลง log
    - ไม่สร้างไฟล์ raw_ai_response.txt
    """
    retries = 0
    wait_time = 1.0
    while retries <= max_retries:
        try:
            response = client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=0.1,
            )
            content = (response.choices[0].message.content or "").strip()
            result = extract_json_from_response(content)
            if result:
                result["_model_used"] = model
                return result
            # ไม่มี JSON ที่ใช้ได้
            logging.warning(f"Model {model} returned non-JSON or invalid JSON.")
            return None

        except Exception as e:
            if is_rate_limit_error(e) and retries < max_retries:
                logging.warning(f"429 from {model}. Retry in {wait_time:.1f}s...")
                time.sleep(wait_time); wait_time *= 2; retries += 1
                continue
            logging.warning(f"Model {model} error: {e}")
            return None

def build_comprehensive_prompt(transcript: str, criteria_data: Dict[str, Any]) -> str:
    """Builds a single, comprehensive prompt for the AI to perform all tasks at once."""
    criteria_ref = []
    for step_no_str, items in criteria_data["criteria_by_step"].items():
        step_no = int(step_no_str)
        for idx, item in enumerate(items, 1):
            name = item.get("name", str(item))
            criteria_ref.append({"id": f"{step_no}.{idx}", "step": step_no, "name": name})

    prompt = f"""
คุณคือผู้ตรวจสอบคุณภาพ (QA) และผู้จัดการรายงานที่เชี่ยวชาญสูง
หน้าที่ของคุณคือทำ 2 อย่างในครั้งเดียว:
1.  วิเคราะห์บทสนทนา (Transcript) อย่างละเอียด และประเมินตามเกณฑ์ที่กำหนด
2.  เขียนสรุปผลการประเมินในเชิงคุณภาพ เป็นภาษาที่สละสลวย

**คุณต้องตอบกลับเป็น JSON object ที่สมบูรณ์เพียง object เดียวเท่านั้น ห้ามมีข้อความอื่นใดนอกเหนือจาก JSON object และห้ามเพิ่ม field อื่นๆ นอกเหนือจากที่กำหนดไว้ในโครงสร้าง JSON ด้านล่างนี้เด็ดขาด**

### 1. บทสนทนาที่ต้องวิเคราะห์ (Transcript)
```text
{transcript}
```

### 2. โครงสร้าง JSON ที่คุณต้องตอบกลับ (Output Format)
คุณต้องกรอกข้อมูลตามโครงสร้างนี้ให้ครบถ้วน โดยอ้างอิงจาก Transcript และเกณฑ์ที่ให้ไว้เท่านั้น:
```json
{{
  "metadata": {{
    "date": "<string or 'ไม่ระบุ'>",
    "branch": "<string or 'ไม่ระบุ'>"
  }},
  "steps": [
    {{
      "step": "<int>",
      "title": "<string>",
      "items": [
        {{
          "id": "<string, e.g., '1.1'>",
          "name": "<string, from CRITERIA_REFERENCE>",
          "status": "พบ|ไม่พบ|พบบางส่วน",
          "reason": "<string, provide a brief reason ONLY if status is 'ไม่พบ' or 'พบบางส่วน'>",
          "evidence": {{
            "exact": "<string, a short quote from the transcript>",
            "sentence": "<string, the full sentence containing the evidence>",
            "offset": ["<string, e.g., '[34..62]'>"]
          }}
        }}
      ]
    }}
  ],
  "summary": {{
    "compliance": {{
      "do": ["<list of observed positive behaviors based on Market Conduct 'Do' list>"],
      "dont": ["<list of observed violations based on Market Conduct 'Don't' list>"]
    }},
    "narrative_reason": "<string, เขียนคำอธิบายเชิงเหตุผล 1 ย่อหน้า ว่าทำไมการสนทนานี้จึงถูกประเมินในระดับคะแนนที่ได้ โดยสังเคราะห์จากภาพรวม ไม่ใช่แค่การลิสต์คะแนน แต่ให้เล่าเรื่องว่าพนักงานทำอะไรได้ดี และขาดตกบกพร่องในขั้นตอนสำคัญใดบ้าง>",
    "strengths": ["<list of 2-3 key strengths>"],
    "improvements": ["<list of 2-3 key areas for improvement>"]
  }}
}}
```

### 3. เกณฑ์การประเมินและข้อมูลอ้างอิง (CRITERIA_REFERENCE)
- **ขั้นตอนและหัวข้อ:** {json.dumps(criteria_data["steps"], ensure_ascii=False)}
- **เกณฑ์ย่อย (ห้ามเพิ่มหรือลบรายการ):** {json.dumps(criteria_ref, ensure_ascii=False, indent=2)}
- **Market Conduct (Do):** {json.dumps(criteria_data.get("compliance_rules", {}).get("Do", []), ensure_ascii=False)}
- **Market Conduct (Don't):** {json.dumps(criteria_data.get("compliance_rules", {}).get("Dont", []), ensure_ascii=False)}

**กฎสำคัญ:**
- **ยึดตาม Transcript เท่านั้น:** ห้ามคาดเดาหรือเสริมข้อมูลที่ไม่มีในบทสนทนา
- **สถานะ:** ทุกเกณฑ์ต้องมีสถานะอย่างใดอย่างหนึ่ง: "พบ", "ไม่พบ", "พบบางส่วน"
- **หลักฐาน (Evidence):** ต้องมาจาก Transcript จริงๆ
- **สรุปผล:** ต้องเขียนสรุป `narrative_reason`, `strengths`, และ `improvements` ให้ครบถ้วนและมีคุณภาพ
- **ห้ามเพิ่ม field อื่นๆ นอกเหนือจากที่กำหนดในโครงสร้าง JSON Output Format เด็ดขาด**
- **หากไม่มีข้อมูลสำหรับ field ใดๆ ให้ใช้ค่าว่าง (empty string for string, empty list for list) แทนการละเว้น field นั้นๆ**
"""
    return prompt

# --- Scoring and Analysis ---

def compute_scores(report: Dict[str, Any]) -> Tuple[Dict[str, Any], Dict[int, Tuple[int, int]]]:
    """Computes scores per step and overall, then adds them to the report."""
    per_step_scores: Dict[int, Tuple[int, int]] = {}
    total_pass = 0
    total_all = 0

    for step in report.get("steps", []):
        step_no = int(step["step"])
        step_pass = 0
        step_all = 0
        for item in step.get("items", []):
            status = (item.get("status") or "").strip()
            step_all += 1
            if status in PASS_STATUSES:
                step_pass += 1
        
        per_step_scores[step_no] = (step_pass, step_all)
        step["score"] = f"{step_pass}/{step_all}"
        total_pass += step_pass
        total_all += step_all

    percent = round(100.0 * total_pass / total_all, 1) if total_all > 0 else 0.0
    
    if "summary" not in report or not isinstance(report["summary"], dict):
        report["summary"] = {}

    report["summary"]["_score_total"] = f"{total_pass}/{total_all}"
    report["summary"]["_score_percent"] = f"{percent}%"
    
    return report, per_step_scores

def normalize_report(report: Dict[str, Any], criteria_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    ทำให้โครงสร้าง report สมบูรณ์เสมอ:
    - ถ้าไม่มี steps: เติมจาก criteria เป็นโครงร่างเริ่มต้น (status ว่าง)
    - เติม summary เปล่า ๆ ถ้าขาด
    """
    if not isinstance(report, dict):
        report = {}

    # สร้าง steps จาก criteria ถ้าไม่มี
    steps_ok = report.get("steps")
    if not (isinstance(steps_ok, list) and len(steps_ok) > 0):
        steps_new = []
        for step_no_str, items in criteria_data.get("criteria_by_step", {}).items():
            step_no = int(step_no_str)
            # หา title จาก criteria_data["steps"] (map หมายเลข -> ชื่อ)
            title = ""
            for s in criteria_data.get("steps", []):
                if int(s.get("step", 0)) == step_no:
                    title = s.get("title", f"ขั้นตอน {step_no}")
                    break
            # สร้าง items ว่าง ๆ ตามเกณฑ์
            new_items = []
            for idx, item in enumerate(items, 1):
                name = item.get("name", str(item))
                new_items.append({
                    "id": f"{step_no}.{idx}",
                    "name": name,
                    "status": "",  # ยังไม่รู้
                    "reason": "",
                    "evidence": {"exact": "", "sentence": "", "offset": []}
                })
            steps_new.append({"step": step_no, "title": title, "items": new_items})
        # sort ตามลำดับขั้น
        steps_new.sort(key=lambda x: int(x["step"]))
        report["steps"] = steps_new

    # ให้มี summary เสมอ
    if not isinstance(report.get("summary"), dict):
        report["summary"] = {
            "compliance": {"do": [], "dont": []},
            "narrative_reason": "",
            "strengths": [],
            "improvements": [],
        }

    # ให้มี metadata เสมอ
    if not isinstance(report.get("metadata"), dict):
        report["metadata"] = {"date": "ไม่ระบุ", "branch": "ไม่ระบุ"}

    return report


def classify_overall(report: Dict[str, Any]) -> str:
    """จัดระดับจากเปอร์เซ็นต์เท่านั้น (ตามเกณฑ์หมายเหตุของผู้ใช้)"""
    summary = report.get("summary", {})
    percent_str = str(summary.get("_score_percent", "0%")).replace('%', '').strip()
    try:
        percent = float(percent_str)
    except Exception:
        percent = 0.0

    # ดีมาก >= 90, ปานกลาง 60–89, ควรปรับปรุง < 60
    if percent >= 90:
        return "ดีมาก"
    elif 60 <= percent < 90:
        return "ปานกลาง"
    else:
        return "ควรปรับปรุง"

# --- DOCX Report Rendering ---

def _apply_run_font(run, *, name="Tahoma", size_pt=11, bold=False):
    """ตั้งฟอนต์/ขนาด/หนา ให้ run รองรับภาษาไทยด้วย eastAsia"""
    run.bold = bold
    run.font.size = Pt(size_pt)
    run.font.name = name
    # สำคัญสำหรับภาษาไทย
    if run._element.rPr is not None:
        rFonts = run._element.rPr.rFonts
        if rFonts is None:
            from docx.oxml import OxmlElement
            rFonts = OxmlElement('w:rFonts')
            run._element.rPr.append(rFonts)
        rFonts.set(qn('w:eastAsia'), name)
        rFonts.set(qn('w:ascii'), name)
        rFonts.set(qn('w:hAnsi'), name)


def _paragraph_with_text(doc, text, *, size_pt=11, bold=False, before=0, after=0, align=WD_ALIGN_PARAGRAPH.LEFT):
    """สร้าง paragraph พร้อมตั้งฟอนต์ตามต้องการ"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.alignment = align
    r = p.add_run(text or "")
    _apply_run_font(r, size_pt=size_pt, bold=bold)
    return p

def _set_cell_text(cell, text, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, size=11):
    """Helper to format text within a table cell (Tahoma 11 by default)."""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(str(text or ""))
    _apply_run_font(run, size_pt=size, bold=bold)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)

# ---------- helpers for layout & font ----------

def _ensure_tahoma_styles(doc, body_size=11, h1_size=16):
    """เซ็ตฟอนต์เริ่มต้นทั้งเอกสารให้เป็น Tahoma (เผื่อภาษาไทย)"""
    st = doc.styles['Normal']
    st.font.name = "Tahoma"
    st.font.size = Pt(body_size)
    st._element.rPr.rFonts.set(qn('w:eastAsia'), "Tahoma")
    st._element.rPr.rFonts.set(qn('w:ascii'), "Tahoma")
    st._element.rPr.rFonts.set(qn('w:hAnsi'), "Tahoma")

    h1 = doc.styles['Heading 1']
    h1.font.name = "Tahoma"
    h1.font.size = Pt(h1_size)
    h1.font.bold = True
    h1._element.rPr.rFonts.set(qn('w:eastAsia'), "Tahoma")
    h1._element.rPr.rFonts.set(qn('w:ascii'), "Tahoma")
    h1._element.rPr.rFonts.set(qn('w:hAnsi'), "Tahoma")

def _printable_width(section):
    # ความกว้างพื้นที่พิมพ์ (นิ้ว) = ความกว้างหน้า - margin ซ้าย/ขวา
    return (section.page_width - section.left_margin - section.right_margin) / Inches(1.0)

def _apply_widths_fit(tbl, widths_in, section, safety=0.88):
    """
    บังคับความกว้างรวมของตารางให้ <= พื้นที่พิมพ์ในหน้า (page_width - margins) * safety
    widths_in: รายการสัดส่วน (หน่วย 'นิ้ว' ที่อยากให้เป็นก่อนสเกล) เช่น [1.2, 0.6, 1.3, 1.5, 3.2, 0.6]
    safety: เผื่อระยะเพื่อไม่ให้ล้น (0.86–0.92)
    """
    page_in = section.page_width.inches
    left_in = section.left_margin.inches
    right_in = section.right_margin.inches

    printable_in = max(0.1, page_in - left_in - right_in) * safety
    sum_desired = sum(widths_in)
    scale = printable_in / sum_desired if sum_desired > 0 else 1.0

    # ตั้งความกว้างรวมตาราง
    tbl.autofit = False
    tbl.allow_autofit = False if hasattr(tbl, "allow_autofit") else None
    tbl.preferred_width = Inches(printable_in)

    # เซ็ตความกว้างแต่ละคอลัมน์
    for j, w in enumerate(widths_in):
        col_w = max(0.2, w * scale)  # กันคอลัมน์เล็กเกิน
        tbl.columns[j].width = Inches(col_w)
        for cell in tbl.columns[j].cells:
            cell.width = Inches(col_w)

def _set_table_cell_margins(table, left=50, right=50, top=30, bottom=30):
    """
    ปรับระยะขอบภายในเซลล์ (หน่วย: twips) ลด padding เพื่อให้ตัวอักษรมีพื้นที่เพิ่ม
    ค่า default ค่อนข้างเล็กอยู่แล้ว แต่อาจช่วยได้มากเมื่อข้อความยาว
    """
    tblPr = table._tbl.tblPr
    tblCellMar = tblPr.xpath("w:tblCellMar")
    if tblCellMar:
        tblCellMar = tblCellMar[0]
    else:
        tblCellMar = OxmlElement('w:tblCellMar')
        tblPr.append(tblCellMar)

    for side, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        elem = tblCellMar.find(qn(f'w:{side}'))
        if elem is None:
            elem = OxmlElement(f'w:{side}')
            tblCellMar.append(elem)
        elem.set(qn('w:w'), str(val))
        elem.set(qn('w:type'), 'dxa')

# ---------- render that APPENDS in LANDSCAPE and fits table ----------

def _repeat_header_on_each_page(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader')
    trPr.append(tblHeader)

def render_report_to_docx(
    original_path: str,
    report: Dict[str, Any],
    per_step_scores: Dict[int, Tuple[int, int]],
    output_file_path: str,
    output_dir: str,
):
    # 1) เปิดไฟล์ input + สไตล์เริ่มต้น
    doc = Document(original_path)
    _ensure_tahoma_styles(doc, body_size=11, h1_size=16)

    # 2) เพิ่ม SECTION ใหม่ (ขึ้นหน้าใหม่) และตั้งเป็น LANDSCAPE (A4)
    new_sec = doc.add_section(WD_SECTION.NEW_PAGE)
    new_sec.orientation = WD_ORIENT.LANDSCAPE
    new_sec.page_width  = Mm(297)   # A4 แนวนอน
    new_sec.page_height = Mm(210)
    new_sec.left_margin   = Inches(0.6)
    new_sec.right_margin  = Inches(0.6)
    new_sec.top_margin    = Inches(0.6)
    new_sec.bottom_margin = Inches(0.6)

    # 3) หัวเรื่อง
    model_name = report.get('_model_used', 'N/A')
    doc.add_paragraph(f"ผลการประเมิน QA โดย AI — Model: {model_name}", style="Heading 1")

    # 4) คำอธิบาย/เมทาดาต้า
    doc.add_paragraph(
        "AI จะทำการวิเคราะห์และให้คะแนนตามเกณฑ์ที่กำหนด\n"
        "โดยคะแนนที่ได้จะถูกคำนวณจากเกณฑ์ทั้งหมด เพื่อให้ผลลัพธ์สะท้อนการทำงานได้อย่างครบถ้วน"
    )
    date_str = report.get('metadata', {}).get('date', 'ไม่ระบุ')
    branch_str = report.get('metadata', {}).get('branch', 'ไม่ระบุ')
    doc.add_paragraph(f"วันที่: {date_str}")
    doc.add_paragraph(f"สาขา: {branch_str}")

    # 5) ตารางรายละเอียดรายขั้นตอน
    steps = report.get("steps", [])
    if not steps:
        doc.add_paragraph()
        doc.add_paragraph("หมายเหตุ: ไม่พบข้อมูลรายละเอียดขั้นตอนจากโมเดล (steps ว่าง)")
    else:
        for step in steps:
            doc.add_paragraph()
            step_no = int(step.get("step", 0))
            title = step.get("title", f"ขั้นตอน {step_no}")
            score = step.get("score", "0/0")

            p = doc.add_paragraph()
            r = p.add_run(f"ขั้นตอน {step_no}) {title} — คะแนน: {score}")
            r.font.name = "Tahoma"; r.font.size = Pt(11); r.bold = True

            _, expected = per_step_scores.get(step_no, (0, 0))
            doc.add_paragraph(f"เกณฑ์ย่อยทั้งหมด: {expected} ข้อ")

            tbl = doc.add_table(rows=1, cols=6)
            tbl.style = "Table Grid"
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            tbl.autofit = False

            # ตารางหลัก 6 คอลัมน์: Criterion, Status, Reason, Exact, Sentence, Offset
            desired = [2.30, 0.60, 1.20, 1.90, 3.00, 0.8]  # รวม ~8.1" ก่อนสเกล
            _apply_widths_fit(tbl, desired, new_sec, safety=0.88)
            _set_table_cell_margins(tbl, left=40, right=40, top=20, bottom=20)

            def _shrink_font(cell, pt=10):
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(pt)

            hdr = tbl.rows[0].cells
            _set_cell_text(hdr[0], "Criterion", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text(hdr[1], "Status",   bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text(hdr[2], "Reason",   bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text(hdr[3], "Exact",    bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text(hdr[4], "Sentence", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text(hdr[5], "Offset",   bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

            # ทำให้หัวตารางซ้ำทุกหน้าเวลา table ยาว
            _repeat_header_on_each_page(tbl.rows[0])

            for it in step.get("items", []):
                row = tbl.add_row().cells
                ev = it.get("evidence", {}) or {}
                _set_cell_text(row[0], it.get("name",""))
                _set_cell_text(row[1], it.get("status",""), align=WD_ALIGN_PARAGRAPH.CENTER)
                _set_cell_text(row[2], it.get("reason",""))
                _set_cell_text(row[3], ev.get("exact",""))
                _set_cell_text(row[4], ev.get("sentence",""))
                off = ev.get("offset", [])
                _set_cell_text(
                    row[5],
                    ", ".join(map(str, off)) if isinstance(off, list) else str(off),
                    align=WD_ALIGN_PARAGRAPH.CENTER
                )

            for row in tbl.rows[1:]:
                # Reason = col 2, Exact = col 3, Sentence = col 4 (หรือ 5 ถ้ามี Offset)
                _shrink_font(row.cells[2], pt=10)
                _shrink_font(row.cells[3], pt=10)
                _shrink_font(row.cells[4], pt=10)

    # 6) สรุปรวม
    doc.add_paragraph()
    p = doc.add_paragraph("สรุปรวม")
    p.style = "Heading 1"

    # บังคับฟอนต์/ขนาด/ตัวหนา + ระยะห่างบน/ล่าง 1 บรรทัด
    for r in p.runs:
        r.font.name = "Tahoma"
        r.font.size = Pt(16)
        r.bold = True

    pf = p.paragraph_format
    pf.space_before = Pt(12)  # ~ 1 บรรทัด
    pf.space_after  = Pt(12)  # ~ 1 บรรทัด

    sumsec = report.get("summary", {}) or {}
    tb = doc.add_table(rows=1, cols=3)
    tb.style = "Table Grid"
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    tb.autofit = False
    _apply_widths_fit(tb, [1.00, 4.00, 1.50], new_sec, safety=0.88)
    _set_table_cell_margins(tb, left=40, right=40, top=20, bottom=20)

    h = tb.rows[0].cells
    _set_cell_text(h[0], "Step", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(h[1], "ชื่อขั้นตอน", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(h[2], "คะแนน", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    for s in steps:
        r = tb.add_row().cells
        _set_cell_text(r[0], s.get("step",""), align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text(r[1], s.get("title",""))
        _set_cell_text(r[2], s.get("score",""), align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()

    p_total = doc.add_paragraph(f"คะแนนรวม: {sumsec.get('_score_total', '0/0')}")
    # ถ้าอยากกำกับให้ไม่ชิดมากเกินไป:
    p_total.paragraph_format.space_before = Pt(0)
    p_total.paragraph_format.space_after  = Pt(6)

    p_pct = doc.add_paragraph(f"คิดเป็นเปอร์เซ็นต์ (%): {sumsec.get('_score_percent', '0%')}")

    level = classify_overall(report)
    pp = doc.add_paragraph(); rr = pp.add_run(f"การประเมินโดยรวม: {level}")
    rr.font.name = "Tahoma"; rr.font.size = Pt(11); rr.bold = True

    reason = (sumsec.get("narrative_reason") or "").strip()
    if reason:
        p = doc.add_paragraph("เหตุผลการจัดระดับ")
        for r in p.runs:
            r.font.name = "Tahoma"; r.font.size = Pt(12); r.bold = True
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)

        p_body = doc.add_paragraph(reason)
        for r in p_body.runs:
            r.font.name = "Tahoma"; r.font.size = Pt(11)

    # จุดเด่น
    strengths = sumsec.get("strengths") or []
    if strengths:
        p = doc.add_paragraph("จุดเด่น")
        for r in p.runs:
            r.font.name = "Tahoma"; r.font.size = Pt(12); r.bold = True
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)

        for s in strengths:
            p_b = doc.add_paragraph(s, style="List Bullet")
            for r in p_b.runs: r.font.name = "Tahoma"; r.font.size = Pt(11)

    # ควรปรับปรุง
    improvements = sumsec.get("improvements") or []
    if improvements:
        p = doc.add_paragraph("ควรปรับปรุง")
        for r in p.runs:
            r.font.name = "Tahoma"; r.font.size = Pt(12); r.bold = True
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)

        for s in improvements:
            p_b = doc.add_paragraph(s, style="List Bullet")
            for r in p_b.runs: r.font.name = "Tahoma"; r.font.size = Pt(11)


    doc.add_paragraph()
    p_note = doc.add_paragraph()
    run_note = p_note.add_run("หมายเหตุ: ตัวอย่างการตัดสินใจของ AI")
    run_note.bold = True
    run_note.font.name = "Tahoma"
    run_note.font.size = Pt(11)

    # ตามด้วย bullet/บรรทัด “ตัวอย่างการตัดสินใจ...” ของคุณเหมือนเดิม
    for line in DECISION_GUIDE_TEXT:
        doc.add_paragraph(f"• {line}")

    # 7) save แบบปลอดภัย
    os.makedirs(output_dir, exist_ok=True)
    tmp = output_file_path + ".tmp"
    doc.save(tmp)
    os.replace(tmp, output_file_path)

# --- Main Execution Logic ---

def main():
    parser = argparse.ArgumentParser(description="AI-Powered QA Summarizer for Call Transcripts")
    parser.add_argument("--product", required=True, help="Product name matching the criteria JSON file (e.g., 'debit_card')")
    parser.add_argument("--criteria-dir", default="criteria", help="Directory containing criteria JSON files")
    parser.add_argument("--input-dir", default="transcript_with_highlight", help="Directory containing input .docx files")
    parser.add_argument("--output-dir", default="transcript_with_highlight_and_ai_summarize", help="Base directory for AI summary output")
    parser.add_argument("--model", default="google/gemma-3-27b-it:free", help="The primary AI model to use")
    parser.add_argument("--fallback-model", default="mistralai/mixtral-8x7b-instruct", help="Fallback model if the primary fails")
    args = parser.parse_args()

    client = get_openai_client()
    models = [args.model, args.fallback_model]
    
    try:
        criteria_data = load_criteria_from_json(args.product, args.criteria_dir)
    except FileNotFoundError as e:
        logging.error(e)
        return

    input_path = os.path.join(os.path.dirname(__file__), args.input_dir, args.product)
    
    output_dir_ai = os.path.join(args.output_dir, f"{args.product}_final_output")
    os.makedirs(output_dir_ai, exist_ok=True)

    files = glob.glob(os.path.join(input_path, "*.docx"))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]

    if not files:
        logging.warning(f"No .docx files found in '{input_path}'.")
        return

    primary_model = args.model
    fallback_model = args.fallback_model
    model_list = [m for m in [primary_model, fallback_model] if m]

    steps = ["Load transcript", "Build prompt", "Call AI", "Score", "Render"]

    with logging_redirect_tqdm():
        total = len(files)
        for idx, docx_path in enumerate(files, start=1):
            fname = os.path.basename(docx_path)

            logging.info("\n" + "-"*100)
            logging.info(f"▶ Processing file {idx}/{total}: {fname}")
            logging.info("-"*100)
            logging.info("-"*100 + "\n")

            t0 = time.time()
            final_output_file_path = None
            skip_current_file = False  # <-- ต้องเริ่มต้นตรงนี้ทุกไฟล์

            steps = ["Load transcript", "Build prompt", "Call AI", "Score", "Render"]
            with tqdm(total=len(steps), desc=f"[{idx}/{total}] {fname}", unit="step", leave=True) as bar:

                # 1) Load transcript
                transcript = read_docx(docx_path)
                bar.update(1)

                if not transcript or not transcript.strip():
                    logging.warning(f"Skipping empty transcript: {fname}")
                    skip_current_file = True

                if not skip_current_file:
                    # 2) Build prompt
                    prompt = build_comprehensive_prompt(transcript, criteria_data)
                    bar.update(1)

                    # 3) Call AI (อัปเดต postfix แสดงรุ่นที่ใช้)
                    report = None
                    for model in models:
                        bar.set_postfix(model=model)
                        report = call_ai_model(
                            client,
                            [{"role": "user", "content": prompt}],
                            model
                        )
                        if report:
                            break
                    bar.update(1)

                    if not report:
                        logging.error(f"All models failed for: {fname}")
                        skip_current_file = True

                if not skip_current_file:
                    # 3.5) Normalize โครงสร้าง report กัน steps ว่าง
                    try:
                        report = normalize_report(report, criteria_data)
                    except Exception as e:
                        logging.error(f"normalize_report failed for {fname}: {e}")
                        skip_current_file = True

                if not skip_current_file:
                    # 4) Score
                    try:
                        scored_report, per_step_scores = compute_scores(report)
                    except Exception as e:
                        logging.error(f"compute_scores failed for {fname}: {e}")
                        skip_current_file = True
                    else:
                        bar.update(1)

                if not skip_current_file:
                    # 5) Render
                    try:
                        output_file_name = f"ai_summary_{fname}"
                        final_output_file_path = os.path.join(output_dir_ai, output_file_name)
                        render_report_to_docx(
                            docx_path,
                            scored_report,
                            per_step_scores,
                            final_output_file_path,
                            output_dir_ai
                        )
                    except Exception as e:
                        logging.error(f"render_report_to_docx failed for {fname}: {e}")
                        skip_current_file = True
                    else:
                        bar.update(1)

                # ถ้าข้ามไฟล์ ให้ดัน progress ไปเต็มแท่งเพื่อจบ bar สวย ๆ
                if skip_current_file:
                    if bar.n < bar.total:
                        bar.update(bar.total - bar.n)

            # ออกจาก with tqdm แล้วค่อยสรุปสถานะไฟล์นี้
            if skip_current_file:
                logging.info(f"⏭ Skipped: {fname}")
            else:
                elapsed = (time.time() - t0) / 60
                # กัน None ถ้าเผื่อไม่มีค่า (ไม่ควรเกิด เพราะ not skip แล้ว)
                final_output_file_path = final_output_file_path or "(no output path)"
                logging.info(
                    f"✅ Completed: {fname} -> {final_output_file_path} | elapsed {elapsed:.2f} min"
                )

if __name__ == "__main__":
    main()
