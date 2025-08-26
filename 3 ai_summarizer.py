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

# Set encoding
try:
    locale.setlocale(locale.LC_ALL, 'th_TH.utf-8')
    logging.info("Successfully set locale to th_TH.utf-8")
except locale.Error as e:
    logging.warning(f"Could not set locale to th_TH.utf-8: {e}")
except Exception as e:
    logging.warning(f"Unexpected error setting locale: {e}")

import docx
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from dotenv import load_dotenv
from openai import OpenAI

# --- Configuration & Constants ---

load_dotenv()
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

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
    """Extracts a JSON object from a string, even if it's embedded in other text."""
    match = re.search(r'\{.*\}', text, re.DOTALL)
    if match:
        json_str = match.group(0)
        try:
            return json.loads(json_str)
        except json.JSONDecodeError as e:
            logging.error(f"Failed to decode JSON from model response. Error: {e}. Raw JSON string: {json_str}")
            return None
    logging.warning(f"No JSON object found in model response. Raw content: {text}")
    return None

def call_ai_model(client: OpenAI, messages: List[Dict[str, Any]], model: str, max_retries: int = 3) -> Optional[Dict[str, Any]]:
    """
    Calls a single AI model with exponential backoff for rate limit errors.
    """
    last_err = None
    retries = 0
    wait_time = 1.0
    while retries <= max_retries:
        try:
            logging.info(f"Attempting completion with model: {model} (Attempt {retries + 1}/{max_retries + 1})")
            response = client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=0.1,
            )
            content = response.choices[0].message.content
            if content:
                logging.info(f"Raw AI Model Response Content: {content}") # Added logging for raw content
                # Save raw content to a file for inspection
                with open("raw_ai_response.txt", "w", encoding="utf-8") as f:
                    f.write(content)
                sys.stdout.flush() # Force flush stdout
                
                result = extract_json_from_response(content)
                if result:
                    result['_model_used'] = model
                    return result
                else:
                    raise ValueError(f"Model response did not contain a valid JSON object. Content: {content}")
            else:
                raise ValueError("Received empty content from API.")
        
        except Exception as e:
            last_err = e
            if is_rate_limit_error(e) and retries < max_retries:
                logging.warning(f"Rate limit exceeded for model {model}. Retrying in {wait_time:.2f} seconds...")
                time.sleep(wait_time)
                wait_time *= 2
                retries += 1
            else:
                # For other errors, we fail immediately and let the fallback handle it.
                logging.warning(f"Failed to get a valid response from model {model}: {e}")
                return None
    
    logging.error(f"All attempts failed for model {model}. Last error: {last_err}")
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


def classify_overall(report: Dict[str, Any]) -> str:
    """Classifies the overall performance based on the score percentage."""
    summary = report.get("summary", {})
    percent_str = summary.get("_score_percent", "0%")
    percent = float(percent_str.replace('%', ''))
    
    if summary.get("compliance", {}).get("dont"):
        return "ควรปรับปรุง"
        
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


def _make_landscape_new_section(doc):
    """
    แทรก section ใหม่เป็นหน้าถัดไป แล้วตั้งแนวนอน (Landscape) ให้ section นั้น
    คืนค่า section ที่สร้าง (ให้เราเขียนรายงานในหน้านี้)
    """
    # เพิ่มหน้าถัดไปเป็น section ใหม่
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    sec.orientation = WD_ORIENT.LANDSCAPE
    # สลับความกว้างสูง
    new_width = sec.page_height
    new_height = sec.page_width
    sec.page_width = new_width
    sec.page_height = new_height
    return sec


def _normalize_existing_big_titles(doc):
    """
    ถ้าในไฟล์เดิมมีพารากราฟชื่อ:
      - Transcript with Highlights
      - Match Summary
    จะตั้งฟอนต์เป็น Tahoma 16 ให้อัตโนมัติ
    """
    targets = {"transcript with highlights", "match summary"}
    for p in doc.paragraphs:
        text = (p.text or "").strip()
        if text.lower() in targets:
            # เคลียร์แล้วเขียนใหม่ด้วย run เดียว (รักษาเนื้อหา)
            p.clear()
            r = p.add_run(text)
            _apply_run_font(r, size_pt=16, bold=True)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT

def _set_cell_text(cell, text, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, size=11):
    """Helper to format text within a table cell (Tahoma 11 by default)."""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(str(text or ""))
    _apply_run_font(run, size_pt=size, bold=bold)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)

def render_report_to_docx(
    original_path: str,
    report: Dict[str, Any],
    per_step_scores: Dict[int, Tuple[int, int]],
    output_file_path: str, # New parameter for output file path
    output_dir: str,       # New parameter for output directory
    insert_page_break: bool = True
):
    """
    เราจะ:
    - เปิดไฟล์เดิม
    - (ปรับหัวข้อ Transcript/Match Summary ในเดิมให้เป็น Tahoma 16 ถ้ามี)
    - แทรก section ใหม่แนวนอน (Landscape)
    - เขียนรายงานด้วย Tahoma 11 (หัวใหญ่ 16)
    """
    doc = docx.Document(original_path)

    # ปรับหัวข้อใหญ่ในหน้าเดิม (ถ้ามี) ให้เป็น Tahoma 16
    _normalize_existing_big_titles(doc)

    # แทรกหน้าใหม่เป็น section แนวนอนสำหรับรายงาน
    if insert_page_break:
        _make_landscape_new_section(doc)
    else:
        # กรณีไม่แทรก section ใหม่ แปลง section ปัจจุบันให้เป็นแนวนอน
        for sec in doc.sections:
            sec.orientation = WD_ORIENT.LANDSCAPE
            new_width = sec.page_height
            new_height = sec.page_width
            sec.page_width = new_width
            sec.page_height = new_height

    # --- 1) หัวรายงานใหญ่ ---
    model_name = report.get('_model_used', 'N/A')
    big_title = f"ผลการประเมิน QA โดย AI — Model: {model_name}"
    _paragraph_with_text(doc, big_title, size_pt=16, bold=True, after=4)

    # บันทึกหมายเหตุการคิดคะแนน/ข้อมูลพื้นฐาน เป็น Tahoma 11
    _paragraph_with_text(doc, SCORING_LOGIC_NOTE, size_pt=11, after=2)

    date_str = report.get('metadata', {}).get('date', 'ไม่ระบุ')
    branch_str = report.get('metadata', {}).get('branch', 'ไม่ระบุ')
    _paragraph_with_text(doc, f"วันที่:  {date_str}", size_pt=11)
    _paragraph_with_text(doc, f"สาขา:  {branch_str}", size_pt=11)

    # --- 2) ตารางรายละเอียดรายขั้นตอน ---
    for step in report.get("steps", []):
        doc.add_paragraph()  # เว้นบรรทัด
        step_no = int(step["step"])
        title = step.get("title", f"ขั้นตอน {step_no}")
        score = step.get("score", "N/A")

        # หัวขั้นตอน (ตัวใหญ่ไม่จำเป็น 16—ใช้ 11 หนา)
        p = doc.add_paragraph()
        r = p.add_run(f"ขั้นตอน {step_no}) {title} — คะแนน: {score}")
        _apply_run_font(r, size_pt=11, bold=True)

        _, expected = per_step_scores.get(step_no, (0, 0))
        _paragraph_with_text(doc, f"เกณฑ์ย่อยทั้งหมด: {expected} ข้อ", size_pt=11, after=2)

        detail_table = doc.add_table(rows=1, cols=6)
        detail_table.style = "Table Grid"
        hdr = detail_table.rows[0].cells
        _set_cell_text(hdr[0], "Criterion", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text(hdr[1], "Status", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text(hdr[2], "Reason", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text(hdr[3], "Exact", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text(hdr[4], "Sentence", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text(hdr[5], "Offset", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

        for item in step.get("items", []):
            row = detail_table.add_row().cells
            evidence = item.get("evidence", {}) or {}
            _set_cell_text(row[0], item.get("name", ""), size=11)
            _set_cell_text(row[1], item.get("status", ""), size=11)
            _set_cell_text(row[2], item.get("reason", ""), size=11)
            _set_cell_text(row[3], evidence.get("exact", ""), size=11)
            _set_cell_text(row[4], evidence.get("sentence", ""), size=11)
            offset_data = evidence.get("offset", [])
            offset_str = ", ".join(map(str, offset_data)) if isinstance(offset_data, list) else str(offset_data)
            _set_cell_text(row[5], offset_str, size=11)

        detail_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row in detail_table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    # --- 3) สรุปรวม ---
    doc.add_paragraph()
    # หัว "สรุปรวม" ให้เด่น (16)
    _paragraph_with_text(doc, "สรุปรวม", size_pt=16, bold=True, after=2)

    summary_data = report.get("summary", {})
    level = classify_overall(report)

    # ตารางสรุปคะแนนรายขั้นตอน
    _paragraph_with_text(doc, "สรุปคะแนนรายขั้นตอน:", size_pt=11, bold=True, after=2)
    score_table = doc.add_table(rows=1, cols=3)
    score_table.style = "Table Grid"
    hdr_cells = score_table.rows[0].cells
    _set_cell_text(hdr_cells[0], "Step", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(hdr_cells[1], "ชื่อขั้นตอน", bold=True)
    _set_cell_text(hdr_cells[2], "คะแนน", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    for step in report.get("steps", []):
        row_cells = score_table.add_row().cells
        _set_cell_text(row_cells[0], step.get("step", ""), align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_text(row_cells[1], step.get("title", ""))
        _set_cell_text(row_cells[2], step.get("score", ""), align=WD_ALIGN_PARAGRAPH.CENTER)

    score_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    doc.add_paragraph()
    _paragraph_with_text(doc, f"คะแนนรวม:  {summary_data.get('_score_total', 'N/A')}", size_pt=11)
    _paragraph_with_text(doc, f"คิดเป็นเปอร์เซ็นต์ (%):  {summary_data.get('_score_percent', 'N/A')}", size_pt=11)
    _paragraph_with_text(doc, f"การประเมินโดยรวม:  {level}", size_pt=11, bold=True)

    # เหตุผล/จุดเด่น/ควรปรับปรุง จาก JSON (ถ้ามี)
    if (narrative := summary_data.get("narrative_reason")):
        p = doc.add_paragraph()
        r1 = p.add_run("เหตุผลประกอบการจัดระดับ:  ")
        _apply_run_font(r1, size_pt=11, bold=True)
        r2 = p.add_run(narrative)
        _apply_run_font(r2, size_pt=11)

    if strengths := summary_data.get("strengths"):
        _paragraph_with_text(doc, "จุดเด่น:", size_pt=11, bold=True)
        for item in strengths:
            _paragraph_with_text(doc, f"• {item}", size_pt=11)

    if improvements := summary_data.get("improvements"):
        _paragraph_with_text(doc, "ควรปรับปรุง:", size_pt=11, bold=True)
        for item in improvements:
            _paragraph_with_text(doc, f"• {item}", size_pt=11)

    doc.add_paragraph()
    # หมายเหตุ (เกณฑ์ตัวอย่าง) — แสดงเป็นคู่มือ ไม่ใช้เป็นเหตุผลจัดระดับ
    _paragraph_with_text(doc, "หมายเหตุ: ตัวอย่างการตัดสินใจของ AI", size_pt=11, bold=True)
    for line in DECISION_GUIDE_TEXT:
        _paragraph_with_text(doc, f"• {line}", size_pt=11)

    os.makedirs(output_dir, exist_ok=True) # Create output directory if it doesn't exist
    doc.save(output_file_path)
    logging.info(f"Successfully rendered report (Landscape + Tahoma) to {output_file_path}")


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

    input_path = os.path.join(os.path.dirname(__file__), args.input_dir, args.product) # Adjusted input path
    output_dir_ai = os.path.join(args.output_dir, f"{args.product}_final_output") # Adjusted output directory for AI summary
    
    files = glob.glob(os.path.join(input_path, "*.docx"))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]

    if not files:
        logging.warning(f"No .docx files found in '{input_path}'.")
        return

    for docx_path in files:
        logging.info(f"--- Processing file: {os.path.basename(docx_path)} ---")
        transcript = read_docx(docx_path)
        if not transcript.strip():
            logging.warning("Skipping empty transcript.")
            continue

        comprehensive_prompt = build_comprehensive_prompt(transcript, criteria_data)
        
        report = None
        for model in models:
            report = call_ai_model(client, [{"role": "user", "content": comprehensive_prompt}], model)
            if report:
                logging.info(f"Successfully received valid report from model: {model}")
                break
            else:
                logging.warning(f"Model {model} failed. Trying next model in the list...")

        if not report:
            logging.error(f"Failed to get a complete report from AI for {docx_path} after trying all models. Skipping.")
            continue

        # --- Scoring and Rendering ---
        scored_report, per_step_scores = compute_scores(report)
        
        # Add a log to inspect the scored_report content
        logging.info(f"Scored Report Content for {os.path.basename(docx_path)}: {json.dumps(scored_report, ensure_ascii=False, indent=2)}")
        
        # Generate new output file path for AI summary
        output_file_name = f"ai_summary_{os.path.basename(docx_path)}"
        final_output_file_path = os.path.join(output_dir_ai, output_file_name)

        render_report_to_docx(docx_path, scored_report, per_step_scores, final_output_file_path, output_dir_ai)
        logging.info(f"--- Finished processing {os.path.basename(docx_path)} ---")


if __name__ == "__main__":
    main()
