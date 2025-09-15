# รันครั้งเดียวให้ครบ: ทำ highlight ทุก transcript ของ product แล้วตามด้วย AI summarizer
# ใช้แบบ:
#   python run_all.py --product debit_card
# หรือ
#   python run_all.py --product telesales

import argparse
import glob
import os
import subprocess
import sys
import time

SEPARATOR = "-" * 100

def find_transcripts(product: str, base_dir: str = "transcript"):
    """หาไฟล์ transcript ภายใต้ transcript/<product>/ รองรับ .txt และ .csv"""
    prod_dir = os.path.join(base_dir, product)
    txts = glob.glob(os.path.join(prod_dir, "*.txt"))
    csvs = glob.glob(os.path.join(prod_dir, "*.csv"))
    files = sorted(txts + csvs)
    return prod_dir, files

def ensure_output_dirs(product: str):
    """สร้างโฟลเดอร์เอาต์พุตที่ใช้ในสองสเต็ปนี้ ถ้ายังไม่มี"""
    highlight_dir = os.path.join("transcript_with_highlight", product)
    final_dir = os.path.join("transcript_with_highlight_and_ai_summarize", f"{product}_final_output")
    os.makedirs(highlight_dir, exist_ok=True)
    os.makedirs(final_dir, exist_ok=True)
    return highlight_dir, final_dir

def run_highlight(product: str, transcript_filename: str) -> bool:
    """
    เรียก 2 keyword_highlight.py สำหรับไฟล์เดียว
    transcript_filename คือชื่อไฟล์อย่างเดียว (เช่น 'call_001.txt')
    """
    cmd = [
        sys.executable,  # ใช้ python interpreter เดียวกับที่รันไฟล์นี้
        "2 keyword_highlight.py",
        "--product", product,
        "--transcript_filename", transcript_filename
    ]
    print(f"  • highlight: {transcript_filename}")
    t0 = time.time()
    proc = subprocess.run(cmd, stdout=sys.stdout, stderr=sys.stderr, text=True)
    ok = (proc.returncode == 0)
    print(f"    -> {'OK' if ok else 'FAILED'} (elapsed {time.time() - t0:.1f}s)")
    return ok

def run_summarizer(product: str, input_dir: str = "transcript_with_highlight") -> bool:
    """
    เรียก 3 ai_summarizer.py ครอบ product เดียว
    หมายเหตุ: โค้ดใน 3 ai_summarizer.py ของคุณควรประกอบ path เป็น <input_dir>/<product>
    ตามที่เคยปรับไว้ก่อนหน้านี้ (input_path = os.path.join(..., args.input_dir, args.product))
    """
    cmd = [
        sys.executable,
        "3 ai_summarizer.py",
        "--product", product,
        "--input-dir", input_dir
    ]
    print(f"▶ summarize: product={product} (input-dir={input_dir})")
    t0 = time.time()
    proc = subprocess.run(cmd, stdout=sys.stdout, stderr=sys.stderr, text=True)
    ok = (proc.returncode == 0)
    print(f"    -> {'OK' if ok else 'FAILED'} (elapsed {(time.time() - t0)/60:.2f} min)")
    return ok

def main():
    parser = argparse.ArgumentParser(description="Run highlight -> summarizer in one go for a single product.")
    parser.add_argument("--product", required=True, help="Product name (e.g., debit_card, telesales)")
    args = parser.parse_args()
    product = args.product

    print("\n" + SEPARATOR)
    print(f"🚀 RUN ALL: product = {product}")
    print(SEPARATOR)

    # 1) เตรียมไฟล์และโฟลเดอร์
    prod_dir, transcripts = find_transcripts(product)
    highlight_dir, final_dir = ensure_output_dirs(product)

    print(f"• transcripts dir : {prod_dir}")
    print(f"• highlight dir   : {highlight_dir}")
    print(f"• final output dir: {final_dir}")
    print(f"• found {len(transcripts)} transcript file(s)")
    print(SEPARATOR)

    if not transcripts:
        print("❌ ไม่พบไฟล์ transcript (*.txt, *.csv) ในโฟลเดอร์นี้")
        sys.exit(1)

    # 2) รัน highlight ทีละไฟล์
    print("STEP 1/2: Highlight transcripts")
    h_ok = True
    for i, path in enumerate(transcripts, start=1):
        fname = os.path.basename(path)
        print(f"[{i}/{len(transcripts)}] {fname}")
        if not run_highlight(product, fname):
            h_ok = False
        print(SEPARATOR)

    # ถ้า highlight มีบางไฟล์พัง ก็ยังไปต่อได้ (เพราะคุณอาจอยากให้ไฟล์ที่สำเร็จ ถูก summarize ต่อ)
    # แต่ถ้าต้องการให้หยุดเลยเมื่อเจอ error ให้ยกเลิกคอมเมนต์สองบรรทัดด้านล่าง:
    # if not h_ok:
    #     print("❌ พบความผิดพลาดตอน highlight. หยุดการรัน.")
    #     sys.exit(1)

    # 3) รัน summarizer ครอบทั้ง product
    print("STEP 2/2: Summarize highlighted docs")
    s_ok = run_summarizer(product, input_dir="transcript_with_highlight")

    print(SEPARATOR)
    if h_ok and s_ok:
        print("✅ All steps completed successfully.")
    else:
        print("⚠ Done with warnings. บางขั้นตอนมีข้อผิดพลาด — โปรดตรวจสอบ log ด้านบน")

if __name__ == "__main__":
    main()
