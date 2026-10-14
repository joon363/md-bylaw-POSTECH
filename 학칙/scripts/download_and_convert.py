#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
POSTECH Regulations (학칙) Downloader & PDF-to-Markdown Converter
----------------------------------------------------------------
1. Downloads regulation PDFs listed in `pdfs.txt` into the `학칙` directory.
2. Converts downloaded PDF files to structured Markdown (.md) format.
"""

import os
import re
import sys
import argparse
import unicodedata
import urllib.request
import urllib.parse
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
import fitz  # PyMuPDF

sys.stdout.reconfigure(encoding='utf-8')

ROOT_DIR = Path(__file__).resolve().parents[2]
PDFS_TXT = ROOT_DIR / "pdfs.txt"
HAKCHIK_DIR = ROOT_DIR / "학칙"
PDF_DIR = HAKCHIK_DIR / "pdf"
MD_DIR = HAKCHIK_DIR / "md"

CIRCLE_MAP = {
    '①': '(1)', '②': '(2)', '③': '(3)', '④': '(4)', '⑤': '(5)',
    '⑥': '(6)', '⑦': '(7)', '⑧': '(8)', '⑨': '(9)', '⑩': '(10)',
    '⑪': '(11)', '⑫': '(12)', '⑬': '(13)', '⑭': '(14)', '⑮': '(15)',
    '⑯': '(16)', '⑰': '(17)', '⑱': '(18)', '⑲': '(19)', '⑳': '(20)'
}

def replace_circles(text: str) -> str:
    for k, v in CIRCLE_MAP.items():
        text = re.sub(re.escape(k) + r'\s*', v + ' ', text)
    return text

def normalize_article_header(line: str) -> str:
    """
    Standardize article headers:
    e.g. 제1조 (목적) ... -> **제1조【목적】** ...
         제20조의2 [제목] ... -> **제20조의2【제목】** ...
         제3조 ... -> **제3조** ...
    """
    # Case with title in parenthesis, brackets or lenticular brackets
    m = re.match(r'^\s*(제\s*\d+\s*조(?:\s*의\s*\d+)*)\s*(?:【(.*?)】|\[(.*?)\]|\((.*?)\))\s*(.*)', line)
    if m:
        art_no = re.sub(r'\s+', '', m.group(1))
        title = (m.group(2) or m.group(3) or m.group(4) or '').strip()
        rest = m.group(5).strip()
        if rest:
            return f'**{art_no}【{title}】** {rest}'
        else:
            return f'**{art_no}【{title}】**'

    # Case without title
    m2 = re.match(r'^\s*(제\s*\d+\s*조(?:\s*의\s*\d+)*)\s+(.*)', line)
    if m2:
        art_no = re.sub(r'\s+', '', m2.group(1))
        rest = m2.group(2).strip()
        return f'**{art_no}** {rest}'

    return line

def sanitize_filename(name: str) -> str:
    # Remove characters forbidden on Windows
    name = re.sub(r'[\\/:*?"<>|]', '_', name)
    # Strip leading/trailing spaces and dots
    return name.strip(' .')

def download_one(url: str, output_dir: Path) -> tuple[bool, str, str]:
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            cd = resp.headers.get('Content-Disposition', '')
            fname = None
            m = re.search(r"filename\*=UTF-8''([^;\s]+)", cd)
            if m:
                fname = urllib.parse.unquote(m.group(1))
            else:
                m2 = re.search(r'filename="?([^";]+)"?', cd)
                if m2:
                    fname = m2.group(1)
            
            if not fname:
                fname = url.rstrip('/').split('/')[-1]
                if not fname.endswith('.pdf'):
                    fname += '.pdf'

            fname = unicodedata.normalize('NFC', fname)
            fname = sanitize_filename(fname)
            if not fname.lower().endswith('.pdf'):
                fname += '.pdf'

            file_path = output_dir / fname
            data = resp.read()
            with open(file_path, 'wb') as f:
                f.write(data)
            return True, url, fname
    except Exception as e:
        return False, url, str(e)

def download_all(urls_file: Path, output_dir: Path, max_workers: int = 10):
    output_dir.mkdir(parents=True, exist_ok=True)
    if not urls_file.exists():
        raise FileNotFoundError(f"URLs file not found: {urls_file}")

    with open(urls_file, 'r', encoding='utf-8') as f:
        urls = [line.strip() for line in f if line.strip()]

    print(f"=== Starting PDF Download ({len(urls)} items) to '{output_dir.name}' ===")
    success_count = 0
    fail_count = 0

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {executor.submit(download_one, url, output_dir): url for url in urls}
        for future in as_completed(future_map):
            ok, url, res = future.result()
            if ok:
                success_count += 1
                print(f"[{success_count}/{len(urls)}] Downloaded: {res}")
            else:
                fail_count += 1
                print(f"[FAIL] {url} -> {res}", file=sys.stderr)

    print(f"=== Download Finished: {success_count} succeeded, {fail_count} failed ===")
    return success_count, fail_count

def extract_page_elements(page: fitz.Page) -> list[tuple[float, str, str]]:
    """
    Extracts text lines and tables from a page in top-to-bottom order.
    Returns list of (y0, element_type, content) where element_type in ('text', 'table').
    """
    elements = []
    
    # 1. Detect tables
    tables = []
    try:
        tabs = page.find_tables()
        tables = tabs.tables if tabs else []
    except Exception:
        tables = []

    table_bboxes = []
    for tab in tables:
        df = tab.extract()
        if not df:
            continue
        table_bboxes.append(tab.bbox)
        
        # Build markdown table
        headers = [str(c or '').strip().replace('\n', ' ') for c in df[0]]
        md_rows = ['| ' + ' | '.join(headers) + ' |', '| ' + ' | '.join(['---'] * len(headers)) + ' |']
        for row in df[1:]:
            vals = [str(c or '').strip().replace('\n', ' ') for c in row]
            md_rows.append('| ' + ' | '.join(vals) + ' |')
        
        elements.append((tab.bbox[1], 'table', '\n'.join(md_rows)))

    # 2. Extract words and cluster by y-coordinate, filtering out tables
    words = page.get_text('words')
    # Filter out words inside table bboxes
    text_words = []
    for w in words:
        # Check if inside any table bbox
        in_tab = False
        for tb in table_bboxes:
            if tb[0] <= w[0] and w[2] <= tb[2] and tb[1] <= w[1] and w[3] <= tb[3]:
                in_tab = True
                break
        if not in_tab:
            # Filter page numbers / header codes
            # e.g. standalone page number at bottom
            if w[1] > 770 and re.match(r'^-?\s*\d+\s*-?$', w[4].strip()):
                continue
            # header doc code at very top
            if w[1] < 70 and re.match(r'^[A-Z]?\d+-\d+(-\d+)?(~\d+)?$', w[4].strip()):
                continue
            text_words.append(w)

    # Group words by block to preserve multi-column or separated blocks
    blocks = {}
    for w in text_words:
        blocks.setdefault(w[5], []).append(w)

    sorted_blocks = sorted(blocks.items(), key=lambda b: (min(w[1] for w in b[1]), min(w[0] for w in b[1])))
    for _, b_words in sorted_blocks:
        b_lines = []
        curr_line = []
        curr_y = None
        for w in sorted(b_words, key=lambda w: (w[1], w[0])):
            if curr_y is None or abs(w[1] - curr_y) < 3.0:
                curr_line.append(w)
                if curr_y is None:
                    curr_y = w[1]
            else:
                b_lines.append(curr_line)
                curr_line = [w]
                curr_y = w[1]
        if curr_line:
            b_lines.append(curr_line)

        for bl in b_lines:
            bl_sorted = sorted(bl, key=lambda w: w[0])
            line_str = " ".join(w[4] for w in bl_sorted).strip()
            # Check standalone page number
            if re.match(r'^-?\s*\d+\s*-?$', line_str):
                continue
            if line_str:
                elements.append((bl_sorted[0][1], 'text', line_str))

    # Sort all elements on page top-to-bottom
    elements.sort(key=lambda el: el[0])
    return elements

def convert_pdf_to_md(pdf_path: Path, md_path: Path):
    doc = fitz.open(pdf_path)
    all_elements = []
    for page in doc:
        all_elements.extend(extract_page_elements(page))

    # Pattern for lines that must begin a new block / paragraph
    new_block_pat = re.compile(
        r'^(-*\s*부\s*칙\s*-*|제\s*\d+\s*[장조절단]|[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]|\d+\.|\(\d+\)|[가-하]\.|\d{4}\s*년|(?:제정|개정)\s*\d+|\[.*?\])'
    )

    blocks = []
    curr_text = ''

    for _, el_type, content in all_elements:
        if el_type == 'table':
            if curr_text:
                blocks.append(('text', curr_text))
                curr_text = ''
            blocks.append(('table', content))
            continue

        s = content.strip()
        if not s:
            if curr_text:
                blocks.append(('text', curr_text))
                curr_text = ''
            continue

        if new_block_pat.match(s):
            if curr_text:
                blocks.append(('text', curr_text))
            curr_text = s
        else:
            if not curr_text:
                curr_text = s
            else:
                if curr_text.endswith(' '):
                    curr_text += s
                else:
                    curr_text += ' ' + s

    if curr_text:
        blocks.append(('text', curr_text))

    md_lines = []
    in_title_zone = True

    for b_type, b in blocks:
        if b_type == 'table':
            md_lines.append("")
            md_lines.append(b)
            md_lines.append("")
            continue

        b = b.strip()
        if not b:
            continue

        # Handle title in the first non-history text block
        if in_title_zone:
            if re.match(r'^\d{4}\s*년', b) or re.match(r'^(제정|개정)\s*\d+', b) or b.startswith('[') or re.match(r'^제\s*\d+\s*[장조]', b):
                in_title_zone = False
            else:
                # Top title heading
                md_lines.append(f"# {b}")
                md_lines.append("")
                in_title_zone = False
                continue

        # Rule 1: 제N장 -> ## 제N장
        m_jang = re.match(r'^제\s*(\d+)\s*[장단]\s*(.*)$', b)
        if m_jang:
            title = re.sub(r'\s+', ' ', m_jang.group(2).strip())
            md_lines.append("")
            md_lines.append(f"## 제{m_jang.group(1)}장 {title}".strip())
            md_lines.append("")
            continue

        # Section: 제N절 -> ### 제N절
        m_jeol = re.match(r'^제\s*(\d+)\s*절\s*(.*)$', b)
        if m_jeol:
            title = re.sub(r'\s+', ' ', m_jeol.group(2).strip())
            md_lines.append("")
            md_lines.append(f"### 제{m_jeol.group(1)}절 {title}".strip())
            md_lines.append("")
            continue

        # Addenda: 부칙 -> ## 부칙
        m_buchik = re.match(r'^-*\s*부\s*칙\s*-*(\s*<.*?>|\s*\[.*?\]|\s*\(.*?\))?$', b)
        if m_buchik:
            md_lines.append("")
            extra = m_buchik.group(1) or ""
            md_lines.append(f"## 부칙{extra}".strip())
            md_lines.append("")
            continue

        # Rule: bold article header **제N조【제목】**
        b = normalize_article_header(b)

        # Rule: circle numbers -> (1), (2), ...
        b = replace_circles(b)

        if b.startswith('**제'):
            md_lines.append("")
            md_lines.append(b)
        elif re.match(r'^\(\d+\)', b):
            md_lines.append("")
            md_lines.append(b)
        else:
            md_lines.append(b)

    # Normalize double blank lines
    cleaned = []
    prev_empty = False
    for line in md_lines:
        line_clean = line.rstrip()
        if not line_clean:
            if not prev_empty:
                cleaned.append("")
                prev_empty = True
        else:
            cleaned.append(line_clean)
            prev_empty = False

    out_content = "\n".join(cleaned).strip() + "\n"
    md_path.write_text(out_content, encoding='utf-8')

def convert_all(pdf_dir: Path, md_dir: Path):
    md_dir.mkdir(parents=True, exist_ok=True)
    pdf_files = sorted(pdf_dir.glob("*.pdf"))
    print(f"=== Starting PDF -> MD Conversion ({len(pdf_files)} files) from '{pdf_dir.name}' to '{md_dir.name}' ===")
    converted = 0
    errors = 0
    for idx, pdf_path in enumerate(pdf_files, 1):
        md_name = unicodedata.normalize('NFC', pdf_path.stem) + ".md"
        md_path = md_dir / md_name
        try:
            convert_pdf_to_md(pdf_path, md_path)
            converted += 1
            if idx % 20 == 0 or idx == len(pdf_files):
                print(f"[{idx}/{len(pdf_files)}] Converted: {md_name}")
        except Exception as e:
            errors += 1
            print(f"[ERROR] Converting {pdf_path.name}: {e}", file=sys.stderr)

    print(f"=== Conversion Completed: {converted} converted, {errors} errors ===")

def main():
    parser = argparse.ArgumentParser(description="Download and convert POSTECH regulations to Markdown.")
    parser.add_argument("--download-only", action="store_true", help="Only download PDFs.")
    parser.add_argument("--convert-only", action="store_true", help="Only convert PDFs to MD.")
    parser.add_argument("--workers", type=int, default=10, help="Number of concurrent download threads.")
    args = parser.parse_args()

    if not args.convert_only:
        download_all(PDFS_TXT, PDF_DIR, max_workers=args.workers)

    if not args.download_only:
        convert_all(PDF_DIR, MD_DIR)

if __name__ == "__main__":
    main()
