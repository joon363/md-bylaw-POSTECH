import os
import re
import sys
import unicodedata
from pathlib import Path
import docx
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.table import Table
import fitz

sys.stdout.reconfigure(encoding='utf-8')

ROOT_DIR = Path(r"c:\Users\joon363\dev\stu40\md-bylaw-POSTECH")

CIRCLE_MAP = {
    '①': '(1)', '②': '(2)', '③': '(3)', '④': '(4)', '⑤': '(5)',
    '⑥': '(6)', '⑦': '(7)', '⑧': '(8)', '⑨': '(9)', '⑩': '(10)',
    '⑪': '(11)', '⑫': '(12)', '⑬': '(13)', '⑭': '(14)', '⑮': '(15)',
    '⑯': '(16)', '⑰': '(17)', '⑱': '(18)', '⑲': '(19)', '⑳': '(20)'
}

def clean_name(s):
    return re.sub(r'\s+', '', unicodedata.normalize('NFC', s))

# Build catalog of all markdown target files and forms
def build_catalog():
    rules_map = {} # normalized_name -> relative_path_from_ROOT_DIR
    forms_map = {}

    # 1. 포항공과대학교 총학생회칙
    dir1 = ROOT_DIR / "포항공과대학교 총학생회칙"
    for f in dir1.glob("*.*"):
        if f.is_dir() or "english" in str(f).lower():
            continue
        if f.suffix in ['.docx', '.pdf']:
            md_name = unicodedata.normalize('NFC', f.stem) + ".md"
            rel = f"포항공과대학교 총학생회칙/{md_name}"
            
            base = unicodedata.normalize('NFC', f.stem)
            rules_map[clean_name(base)] = rel
            
            # Extract main rule name e.g. "학생총투표규칙" from "의결기구규칙 제1호 - 학생총투표규칙 (2023.12.24. 개정)"
            m = re.search(r'-\s*([가-힣\s]+?)(?:\s*\(|$)', base)
            if m:
                rules_map[clean_name(m.group(1))] = rel
            if "총학생회칙" in base:
                rules_map[clean_name("총학생회칙")] = rel
                rules_map[clean_name("포항공과대학교 총학생회칙")] = rel

    # Specific common aliases
    alias_dict = {
        '전체학생대의원규칙': '의결기구규칙 제2호 - 전체학생대의원회의규칙 (2023.12.24. 개정).md',
        '회의진행세칙': '의결기구규칙 제5호 - 회의진행규칙 (2023.12.24. 개정).md',
        '기록물관리세칙': '일반규칙 제1호 - 기록물관리규칙 (2023.12.24. 개정).md',
        '기록물관리규칙': '일반규칙 제1호 - 기록물관리규칙 (2023.12.24. 개정).md',
    }
    for a_name, a_file in alias_dict.items():
        rules_map[clean_name(a_name)] = f"포항공과대학교 총학생회칙/{a_file}"

    # 2. 기구별 자치규칙
    dir2 = ROOT_DIR / "기구별 자치규칙"
    for f in dir2.rglob("*.*"):
        if f.is_dir():
            continue
        if f.suffix == '.pdf' or (f.suffix == '.docx' and '생활관자치회' in str(f)):
            md_name = unicodedata.normalize('NFC', f.stem) + ".md"
            rel_dir = f.parent.relative_to(ROOT_DIR).as_posix()
            rel = f"{rel_dir}/{md_name}"
            
            base = unicodedata.normalize('NFC', f.stem)
            rules_map[clean_name(base)] = rel
            
            # Additional clean matches
            m_dept = re.search(r'^\d+\.\s*([가-힣\s]+?)(?:\s*\(|$|_)', base)
            if m_dept:
                rules_map[clean_name(m_dept.group(1))] = rel
                if "자치규칙" not in m_dept.group(1):
                    rules_map[clean_name(m_dept.group(1) + "자치규칙")] = rel
                    rules_map[clean_name(m_dept.group(1) + "학생회칙")] = rel

            # Keyword matching
            for kw in ['동아리등록금세칙', '동아리방배정평가세칙', '동아리연합회자치규칙', '도서관자치위원회자치규칙',
                       '생활관자치회칙', '총여학생회자치규칙', '방송국자치규칙', '학생·소수자인권위원회자치규칙',
                       '학생소수자인권위원회자치규칙', '학생교육위원회자치규칙', '행사지원위원회자치규칙',
                       '무은재학생회내규', '무은재학생회자치규칙']:
                if kw in clean_name(base):
                    rules_map[clean_name(kw)] = rel

    # 3. 총학생회 서류 양식
    dir3 = ROOT_DIR / "총학생회 서류 양식"
    for f in dir3.glob("*.*"):
        if f.is_dir():
            continue
        if f.suffix in ['.docx', '.xlsx']:
            base = unicodedata.normalize('NFC', f.name)
            rel = f"총학생회 서류 양식/{base}"
            m = re.search(r'\]\s*([가-힣\s\d]+?)(?:\s*\(|$)', base)
            if m:
                form_name = clean_name(m.group(1))
                if form_name not in forms_map or f.suffix == '.docx':
                    forms_map[form_name] = rel

    return rules_map, forms_map

RULES_MAP, FORMS_MAP = build_catalog()

def replace_circles(text):
    for k, v in CIRCLE_MAP.items():
        text = re.sub(k + r'\s*', v + ' ', text)
    return text

def normalize_article_header(line):
    # Only if line starts with 제N조 (article header definition)
    # Formats to bold with bold brackets: **제N조【제목】**
    m = re.match(r'^\s*(제\s*\d+\s*조(?:\s*의\s*\d+)?)\s*(?:【(.*?)】|\[(.*?)\]|\(([가-힣\s,·/]+?)\))\s*(.*)', line)
    if m:
        art_no = re.sub(r'\s+', '', m.group(1))
        title = (m.group(2) or m.group(3) or m.group(4) or '').strip()
        rest = m.group(5)
        if rest:
            return f'**{art_no}【{title}】** {rest}'
        else:
            return f'**{art_no}【{title}】**'

    # If line starts with 제N조 without title
    m2 = re.match(r'^\s*(제\s*\d+\s*조(?:\s*의\s*\d+)?)\s+(.*)', line)
    if m2:
        art_no = re.sub(r'\s+', '', m2.group(1))
        rest = m2.group(2)
        return f'**{art_no}** {rest}'

    return line

def link_rules_and_forms(text, current_file_path):
    current_dir = (ROOT_DIR / current_file_path).parent

    def repl_quote(match):
        inner = match.group(1).strip()
        c_inner = clean_name(inner)

        target = None
        if c_inner in RULES_MAP:
            target = RULES_MAP[c_inner]
        else:
            for k, v in RULES_MAP.items():
                if k and (k in c_inner or c_inner in k) and len(k) >= 4:
                    target = v
                    break

        if not target:
            if c_inner in FORMS_MAP:
                target = FORMS_MAP[c_inner]
            else:
                for k, v in FORMS_MAP.items():
                    if k and (k in c_inner or c_inner in k) and len(k) >= 4:
                        target = v
                        break

        if target:
            abs_target = ROOT_DIR / target
            rel = os.path.relpath(abs_target, current_dir).replace('\\', '/')
            if not rel.startswith('.') and not rel.startswith('/'):
                rel = './' + rel
            encoded = rel.replace(' ', '%20')
            return f'[{inner}]({encoded})'
        
        return match.group(0)

    # 1. Replace 「...」
    text = re.sub(r'「([^」]+)」', repl_quote, text)

    # 2. Replace [...] only when it matches a known rule or form, and NOT immediately preceded by 제N조
    def repl_bracket(match):
        inner = match.group(1).strip()
        c_inner = clean_name(inner)

        # Do NOT replace if preceded by 제N조 (it is an article title)
        start = match.start()
        prefix = text[:start]
        if re.search(r'제\s*\d+\s*조\s*$', prefix):
            return match.group(0)

        # Skip generic terms
        if c_inner in ['목적', '정의', '구성', '명칭', '소재', '의무', '권리', '회원', '임기', '자격', '원칙',
                       '기본이념', '단위', '적용', '업무', '권한', '소집', '자문', '안건', '전문', '활동', '기구']:
            return match.group(0)
        if re.match(r'^(일반규칙|의결기구규칙|자치기구규칙|재정운용규칙|총학생회세칙|특별기구규칙)\s*제\d+호', c_inner):
            return match.group(0)
        
        target = None
        if c_inner in RULES_MAP:
            target = RULES_MAP[c_inner]
        elif c_inner in FORMS_MAP:
            target = FORMS_MAP[c_inner]

        if target:
            abs_target = ROOT_DIR / target
            rel = os.path.relpath(abs_target, current_dir).replace('\\', '/')
            if not rel.startswith('.') and not rel.startswith('/'):
                rel = './' + rel
            encoded = rel.replace(' ', '%20')
            return f'[{inner}]({encoded})'

        return match.group(0)

    text = re.sub(r'\[([^\]]+)\](?!\()', repl_bracket, text)
    return text

def is_toc_line(line):
    if re.search(r'[\.·•]{3,}', line):
        return True
    if re.search(r'[\.·\s]{2,}\d+(~\d+)?\s*$', line):
        return True
    return False

def convert_docx_file(docx_path, out_md_path, rel_path):
    doc = docx.Document(docx_path)
    md_lines = []
    in_title_zone = True

    for child in doc.element.body:
        if isinstance(child, CT_P):
            p = Paragraph(child, doc)
            text = p.text.strip()
            if not text:
                continue

            sublines = [sl.strip() for sl in text.split('\n') if sl.strip()]
            for sl in sublines:
                # Title detection at top
                if in_title_zone:
                    if re.match(r'^\d{4}\s*년', sl) or sl.startswith('['):
                        in_title_zone = False
                    else:
                        md_lines.append(f"# {sl}")
                        md_lines.append("")
                        in_title_zone = False
                        continue

                # Rule: bold article headers with bold brackets **제N조【제목】**
                sl = normalize_article_header(sl)

                # Rule 2: circles -> (1), (2)
                sl = replace_circles(sl)

                # Rule 1: 제N장 -> ##
                m_jang = re.match(r'^제\s*(\d+)\s*[장단]\s*(.*)$', sl)
                if m_jang:
                    title = re.sub(r'\s+', ' ', m_jang.group(2).strip())
                    md_lines.append("")
                    md_lines.append(f"## 제{m_jang.group(1)}장 {title}".strip())
                    md_lines.append("")
                    continue

                # Section: 제N절 -> ###
                m_jeol = re.match(r'^제\s*(\d+)\s*절\s*(.*)$', sl)
                if m_jeol:
                    title = re.sub(r'\s+', ' ', m_jeol.group(2).strip())
                    md_lines.append("")
                    md_lines.append(f"### 제{m_jeol.group(1)}절 {title}".strip())
                    md_lines.append("")
                    continue

                # Addenda: 부칙 -> ## 부칙
                m_buchik = re.match(r'^-*\s*부\s*칙\s*-*(\s*<.*?>|\s*\[.*?\])?$', sl)
                if m_buchik:
                    md_lines.append("")
                    extra = m_buchik.group(1) or ""
                    md_lines.append(f"## 부칙{extra}".strip())
                    md_lines.append("")
                    continue

                # Rule 5: link rules and forms
                sl = link_rules_and_forms(sl, rel_path)

                if sl.startswith('**제'):
                    md_lines.append("")
                    md_lines.append(sl)
                elif re.match(r'^\(\d+\)', sl):
                    md_lines.append("")
                    md_lines.append(sl)
                else:
                    md_lines.append(sl)

        elif isinstance(child, CT_Tbl):
            table = Table(child, doc)
            md_lines.append("")
            rows = table.rows
            if rows:
                headers = [cell.text.strip().replace('\n', ' ') for cell in rows[0].cells]
                md_lines.append("| " + " | ".join(headers) + " |")
                md_lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
                for row in rows[1:]:
                    vals = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
                    md_lines.append("| " + " | ".join(vals) + " |")
            md_lines.append("")

    # Clean double empty lines
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
    out_md_path.write_text(out_content, encoding='utf-8')
    print(f"[DOCX -> MD] Successfully converted: {out_md_path.name}")

def convert_pdf_file(pdf_path, out_md_path, rel_path):
    doc = fitz.open(pdf_path)
    all_raw_lines = []

    is_dongari_sechik2 = "동아리방 배정 평가 세칙" in pdf_path.name

    for page_idx, page in enumerate(doc):
        t = page.get_text('text')
        lines = [l for l in t.split('\n')]
        
        # Filter standalone page numbers
        non_empty = [i for i, l in enumerate(lines) if l.strip()]
        if non_empty and re.match(r'^-?\s*\d+\s*-?$', lines[non_empty[0]].strip()):
            lines[non_empty[0]] = ''
        if non_empty and re.match(r'^-?\s*\d+\s*-?$', lines[non_empty[-1]].strip()):
            lines[non_empty[-1]] = ''
        all_raw_lines.extend(lines)

    # Detect TOC boundary if present
    toc_start = None
    toc_end = None
    for i, l in enumerate(all_raw_lines):
        s = l.strip()
        if is_toc_line(s) and toc_start is None:
            toc_start = i
        if toc_start is not None and i > toc_start:
            if re.match(r'^제\s*1\s*장', s) and not is_toc_line(s):
                toc_end = i
                break
            if re.match(r'^제\s*1\s*조', s):
                toc_end = i
                break

    # Merge lines into logical blocks
    blocks = []
    curr = ''
    new_block_pat = re.compile(
        r'^(-*\s*부\s*칙\s*-*|제\s*\d+\s*[장조절단]|[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]|\d+\.|\(\d+\)|[가-하]\.|\d{4}\s*년|\[.*?\])'
    )

    in_toc = False
    for i, l in enumerate(all_raw_lines):
        s = l.strip()
        if not s:
            if curr:
                blocks.append((curr, in_toc))
                curr = ''
            continue

        if toc_start is not None and i >= toc_start and (toc_end is None or i < toc_end):
            in_toc = True
        else:
            in_toc = False

        if is_toc_line(s) or in_toc:
            if curr:
                blocks.append((curr, False))
                curr = ''
            blocks.append((s, True))
            continue

        if new_block_pat.match(s):
            if curr:
                blocks.append((curr, False))
            curr = l
        else:
            if not curr:
                curr = l
            else:
                if curr.endswith(' '):
                    curr = curr + l.lstrip()
                else:
                    curr = curr + l

    if curr:
        blocks.append((curr, in_toc))

    md_lines = []
    in_title_zone = True

    for b, was_toc in blocks:
        b = b.strip()
        if not b:
            continue

        if was_toc:
            md_lines.append(b)
            continue

        # Handle title at start
        if in_title_zone:
            if re.match(r'^\d{4}\s*년', b) or b.startswith('[') or re.match(r'^제\s*\d+\s*[장조]', b) or is_toc_line(b):
                in_title_zone = False
            else:
                md_lines.append(f"# {b}")
                md_lines.append("")
                in_title_zone = False
                continue

        # Rule: bold article headers with bold brackets **제N조【제목】**
        b = normalize_article_header(b)

        # Rule 2: circles -> (1), (2)
        b = replace_circles(b)

        # Rule 1: 제N장 -> ##
        m_jang = re.match(r'^제\s*(\d+)\s*[장단]\s*(.*)$', b)
        if m_jang:
            title = re.sub(r'\s+', ' ', m_jang.group(2).strip())
            md_lines.append("")
            md_lines.append(f"## 제{m_jang.group(1)}장 {title}".strip())
            md_lines.append("")
            continue

        # Section: 제N절 -> ###
        m_jeol = re.match(r'^제\s*(\d+)\s*절\s*(.*)$', b)
        if m_jeol:
            title = re.sub(r'\s+', ' ', m_jeol.group(2).strip())
            md_lines.append("")
            md_lines.append(f"### 제{m_jeol.group(1)}절 {title}".strip())
            md_lines.append("")
            continue

        # Addenda: 부칙 -> ## 부칙
        m_buchik = re.match(r'^-*\s*부\s*칙\s*-*(\s*<.*?>|\s*\[.*?\])?$', b)
        if m_buchik:
            md_lines.append("")
            extra = m_buchik.group(1) or ""
            md_lines.append(f"## 부칙{extra}".strip())
            md_lines.append("")
            continue

        # Rule 5: link rules and forms
        b = link_rules_and_forms(b, rel_path)

        if b.startswith('**제'):
            md_lines.append("")
            md_lines.append(b)
        elif re.match(r'^\(\d+\)', b):
            md_lines.append("")
            md_lines.append(b)
        else:
            md_lines.append(b)

    # Special handling: if dongari sechik2, format table nicely
    if is_dongari_sechik2:
        full_text = "\n".join(md_lines)
        if "항목" in full_text and "평가 요소" in full_text:
            table_md = """
| 항목 | 평가 요소 | 배점 | 비고 |
| --- | --- | --- | --- |
| 1 | 물품 보관 및 사용 | 동아리 활동에 필요한 물품의 종류와 개수(10), 그리고 이 물품 보관과 사용의 어려움(10)을 종합적으로 평가 | 20 | 기존에 동아리방이 없는 경우는 만점으로 평가. 물품은 개인 소유, 음식물 및 주류, 화기, 침구류를 취급하지 않는다. |
| 2 | 기존 동아리 방 사용 문제 와 새 동방 활용 계획 | 기존 동아리방에서 활동을 하며 겪는 어려움(10)과 새 동방 배정 시 활용 계획(10)을 종합적으로 평가 | 20 | 기존에 동아리방이 없는 경우는 어려움 항목을 만점으로 평가 함. 내부 회의는 고려하지 않는다. 단, 물품 보관과 3번 항목(크기, 위치, 적합성)에 관련한 항목은 평가하지 않는다. |
| 3 | 동아리방 적합성 | 배정받을 동아리방의 크기(10), 위치(10) 그리고 주변환경과 동아리와의 적합성(10)을 종합적으로 평가 | 30 | 주변 동아리 및 다른 단체나 교내 구성원에 피해를 주게 될 것으로 보이는 경우 크게 감점한다. |
| 4 | 사회 기여도 | 동아리 활동으로 학교 밖 사회에서 봉사, 공연, 세미나 및 기타 활동을 종합적으로 평가 | 5 | 최근 1년 내 동아리 자체 혹은 대표로 구성원 전체 혹은 일부 구성원이 참여한 활동만 인정 |
| 5 | 내부 활성도 | 개강총회, 종강총회 등 동아리 목적에 맞는 내부 활동(10) | 10 | 최근 1년 내 활동만 평가 |
| 6 | 징계 | 1년간 경고 및 징계 기록이 없는 경우 15점을 부여하고, 경고 1회는 10점, 경고 3회 이상은 0점을 부여한다. | 15 | 동아리연합회장이 공개하여 평가에 반영하도록 권고 |
"""
            m_tbl = re.search(r'(항목\s*\n\s*평가 요소[\s\S]*?)(?=\*\*제10조|\Z)', full_text)
            if m_tbl:
                full_text = full_text[:m_tbl.start()] + table_md.strip() + "\n\n" + full_text[m_tbl.end():]
                md_lines = full_text.split('\n')

    # Clean double empty lines
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
    out_md_path.write_text(out_content, encoding='utf-8')
    print(f"[PDF -> MD] Successfully converted: {out_md_path.name}")

def main():
    print("=== Starting Conversion ===")
    
    # 1. 포항공과대학교 총학생회칙 (except english)
    dir1 = ROOT_DIR / "포항공과대학교 총학생회칙"
    docx_dir = dir1 / "docx"
    pdf_dir = dir1 / "pdf"
    md_dir = dir1 / "md"
    md_dir.mkdir(exist_ok=True)

    for f in sorted(docx_dir.glob("*.docx")):
        rel_path = f"포항공과대학교 총학생회칙/md/{f.stem}.md"
        out_md = md_dir / f"{f.stem}.md"
        convert_docx_file(f, out_md, rel_path)

    for f in sorted(pdf_dir.glob("*.pdf")):
        rel_path = f"포항공과대학교 총학생회칙/md/{f.stem}.md"
        out_md = md_dir / f"{f.stem}.md"
        convert_pdf_file(f, out_md, rel_path)

    # 2. 기구별 자치규칙 (all PDFs + 생활관자치회 docx)
    dir2 = ROOT_DIR / "기구별 자치규칙"
    for f in sorted(dir2.rglob("*.*")):
        if f.is_dir():
            continue
        norm_stem = unicodedata.normalize('NFC', f.stem)
        rel_dir = f.parent.relative_to(ROOT_DIR).as_posix()
        rel_path = f"{rel_dir}/{norm_stem}.md"
        out_md = f.parent / f"{norm_stem}.md"
        
        if f.suffix == '.pdf':
            convert_pdf_file(f, out_md, rel_path)
        elif f.suffix == '.docx' and '생활관자치회' in str(f):
            convert_docx_file(f, out_md, rel_path)

    print("=== All Conversions Completed Successfully! ===")

if __name__ == '__main__':
    main()
