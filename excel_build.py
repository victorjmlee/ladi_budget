# -*- coding: utf-8 -*-
"""
엑셀용 2차원 배열(aoa) 생성. 코드 조회 시 블록 follow 라인(페이지 넘김 포함) 처리.
"""
import re
from typing import List, Tuple, Optional

from parser import (
    EXCEL_HEADER,
    line_to_excel_rows,
    take_last_amount_from_line,
    take_first_amount_from_line,
    extract_pdf_meta,
    apply_default_meta,
    is_3digit_start_line,
)
from pdf_extract import (
    get_page_lines,
    get_next_lines_until_next_block,
    get_line_index_in_page,
    get_block_start,
)


def _norm(s: str) -> str:
    return (s or "").replace(" ", "").replace("ㅇ", "○")


def _only_number(t: str) -> bool:
    """금액만 있는 줄. 01·02 등 두 자리 편성목 코드는 제외."""
    s = (t or "").strip()
    if not s:
        return False
    if s == "0":
        return True
    if re.match(r"^\d{3,6}$", s):
        return True
    if re.match(r"^\d{1,3}(?:,\d{3})+$", s):
        return True
    if re.match(r"^△\d[\d,]*$", s):
        return True
    return False


def _is_amount(s: str) -> bool:
    t = (s or "").strip()
    if not t:
        return False
    if re.match(r"^\d{1,3}(?:,\d{3})+$", t):
        return True
    if re.match(r"^△\d[\d,]*$", t):
        return True
    if re.match(r"^△?\d{1,6}$", t):
        return True
    return False


def _parse_triple_amounts(line: str):
    """줄이 금액만 2~3개(라벨 없음)이면 (예산, 전년도, 비교증감) 반환. 한글이 있으면 None."""
    t = (line or "").strip()
    if not t or re.search(r"[가-힣]", t):
        return None
    for split_re in (r"\s{2,}", r"\s+"):
        parts = re.split(split_re, t)
        parts = [p.replace(" ", "").strip() for p in parts if p.strip()]
        if len(parts) == 2 and _is_amount(parts[0]) and _is_amount(parts[1]):
            return (parts[0], parts[1], "")
        if len(parts) >= 3:
            a, b, c = parts[-3], parts[-2], parts[-1]
            if _is_amount(a) and _is_amount(b) and (_is_amount(c) or c.startswith("△")):
                return (a, b, c)
    return None


def _is_formula_amount_line(t: str) -> bool:
    """식+금액만 있는 줄 (100,000원*11명*2회  2,200). 01·02, -로 시작하는 항목은 제외."""
    t = (t or "").strip()
    if not t:
        return False
    if re.match(r"^[가-힣▣■○ㅇ\-]", t):
        return False
    if re.match(r"^\d{2}(?:\s|$|[가-힣])", t):
        return False
    return True


def _next_is_label_row(next_line: Optional[str]) -> bool:
    if not next_line:
        return False
    t = next_line.strip()
    return bool(re.match(r"^[가-힣▣■○ㅇ\-]", t)) or bool(re.match(r"^\d{2}\s", t))


def _is_page_header(line: str) -> bool:
    """페이지 헤더/테이블 헤더만 제외. 기도시는 201 블록 안에서는 포함됨."""
    t = (line or "").strip()
    if not t:
        return True
    if re.match(r"^부서\s*[:：]", t) or re.match(r"^정책\s*[:：]", t) or re.match(r"^단위\s*[:：]", t):
        return True
    if re.match(r"^\(단위\s*[:：]?\s*천원\)", t) or "(단위:천원)" in t:
        return True
    # 테이블 헤더: "부서ㆍ정책ㆍ단위(회계)ㆍ세부사업ㆍ편성목 예산액 전년도 예산액 비교증감"
    if "부서ㆍ정책ㆍ단위" in t and "편성목" in t:
        return True
    if "세부사업" in t and "편성목" in t and "예산액" in t:
        return True
    if t in ("예산액", "전년도", "비교증감"):
        return True
    if "예산액" in t and "전년도" in t and "비교증감" in t:
        return True
    return False


def build_excel_aoa(
    export_lines_with_pages: List[Tuple[str, int, int]],
    page_texts: List[str],
    doc_meta: Optional[dict] = None,
) -> List[List[str]]:
    """
    export_lines_with_pages: [(줄텍스트, 페이지인덱스, 줄인덱스), ...]
    page_texts: 페이지별 전체 텍스트 (정규화된 것)
    doc_meta: 문서 첫 페이지 등에서 추출한 부서/정책/단위 (fallback)
    """
    aoa: List[List[str]] = [EXCEL_HEADER]
    export_norm = [_norm(l) for l, _, _ in export_lines_with_pages]
    page_first_row: dict = {}
    sheet_meta_filled = False
    doc_meta = doc_meta or {}
    if not doc_meta and page_texts:
        doc_meta = apply_default_meta(extract_pdf_meta(page_texts[0]), {})

    def export_contains(page_line: str) -> bool:
        n = _norm(page_line)
        if n in export_norm:
            return True
        for ex in export_norm:
            if ex in n or n in ex:
                return True
        return False

    j = 0
    last_processed = None  # (norm_line, page_index, line_index) - 바로 직전 처리한 것만 체크
    
    while j < len(export_lines_with_pages):
        line, page_index, line_index = export_lines_with_pages[j]
        
        # 바로 직전과 같은 줄(같은 페이지, 같은 line_index)이면 스킵 (연속 중복 제거)
        current_key = (_norm(line), page_index, line_index)
        if last_processed and last_processed == current_key:
            j += 1
            continue
        last_processed = current_key
        page_lines = get_page_lines(page_texts[page_index]) if page_index < len(page_texts) else []
        # line_index를 이미 알고 있으므로 get_line_index_in_page 호출 불필요
        block_start = get_block_start(page_lines, line_index) if page_lines else -1
        next_line_in_page = None
        if 0 <= line_index < len(page_lines) - 1:
            next_line_in_page = page_lines[line_index + 1]

        page_meta = extract_pdf_meta(page_texts[page_index]) if page_index < len(page_texts) else {}
        page_meta = apply_default_meta(page_meta, doc_meta)

        if line_index >= 0 and block_start >= 0 and block_start != line_index:
            block_start_line = page_lines[block_start] if block_start < len(page_lines) else ""
            if block_start_line and export_contains(block_start_line):
                j += 1
                continue

        row_list = line_to_excel_rows(line)
        used_next_export_line = False
        for row in row_list:
            next_ln = next_line_in_page or (export_lines_with_pages[j + 1][0] if j + 1 < len(export_lines_with_pages) else None)
            bullet = bool(row[3] and row[3].strip() and row[3].strip()[0] in "○ㅇ■▣-")
            if bullet and not row[4] and next_ln:
                amt = take_first_amount_from_line(next_ln)
                if amt:
                    row[4] = amt
                    if j + 1 < len(export_lines_with_pages) and next_ln == export_lines_with_pages[j + 1][0]:
                        used_next_export_line = True
            if row[3] and row[3] != "편성목":
                if not sheet_meta_filled:
                    m = apply_default_meta(doc_meta, {})
                    row[0], row[1], row[2] = m.get("부서", ""), m.get("정책", ""), m.get("단위", "")
                    sheet_meta_filled = True
                elif page_index not in page_first_row:
                    row[0], row[1], row[2] = page_meta.get("부서", ""), page_meta.get("정책", ""), page_meta.get("단위", "")
                    page_first_row[page_index] = True
            aoa.append(row)

        is_block_start = is_3digit_start_line(line) and (block_start < 0 or line_index == block_start)
        if is_block_start and page_index < len(page_texts):
            follow_lines = get_next_lines_until_next_block(page_texts, page_index, line_index)
            used_as_amount_only = set()
            last_row_with_no_amount = None
            if len(aoa) >= 2 and aoa[-1][3] != "편성목" and not aoa[-1][4]:
                last_row_with_no_amount = aoa[-1]
            elif len(aoa) >= 3 and aoa[-1][3] == "편성목" and aoa[-2][3] and not aoa[-2][4]:
                last_row_with_no_amount = aoa[-2]
            for fi, follow_line in enumerate(follow_lines):
                if fi in used_as_amount_only:
                    continue
                t = follow_line.strip()
                triple = _parse_triple_amounts(follow_line)
                if triple is not None and len(aoa) >= 2:
                    prev = aoa[-1]
                    if prev[3] == "편성목" and len(aoa) >= 3:
                        prev = aoa[-2]
                    if prev[3] and prev[3] != "편성목":
                        if not prev[4]:
                            prev[4] = triple[0]
                        if triple[1] and not prev[5]:
                            prev[5] = triple[1]
                        if triple[2] and not prev[6]:
                            prev[6] = triple[2]
                    used_as_amount_only.add(fi)
                    continue
                if _only_number(t):
                    n1 = follow_lines[fi + 1].strip() if fi + 1 < len(follow_lines) else ""
                    n2 = follow_lines[fi + 2].strip() if fi + 2 < len(follow_lines) else ""
                    if last_row_with_no_amount is not None and _only_number(n1) and _only_number(n2):
                        last_row_with_no_amount[4] = t
                        last_row_with_no_amount[5] = n1
                        last_row_with_no_amount[6] = n2
                        last_row_with_no_amount = None
                        used_as_amount_only.add(fi)
                        used_as_amount_only.add(fi + 1)
                        used_as_amount_only.add(fi + 2)
                        continue
                    if last_row_with_no_amount is not None:
                        last_row_with_no_amount[4] = t
                        last_row_with_no_amount = None
                    continue
                next_in_block = follow_lines[fi + 1] if fi + 1 < len(follow_lines) else None
                if _is_formula_amount_line(t):
                    amt = take_last_amount_from_line(follow_line)
                    if amt and last_row_with_no_amount is not None:
                        last_row_with_no_amount[4] = amt
                        last_row_with_no_amount = None
                        used_as_amount_only.add(fi)
                        continue
                    if last_row_with_no_amount is not None and next_in_block is not None and _only_number(next_in_block.strip()):
                        last_row_with_no_amount[4] = next_in_block.strip()
                        last_row_with_no_amount = None
                        used_as_amount_only.add(fi)
                        used_as_amount_only.add(fi + 1)
                        continue
                if _is_page_header(follow_line):
                    continue
                if _is_formula_amount_line(t) and not take_last_amount_from_line(follow_line):
                    continue
                sub_list = line_to_excel_rows(follow_line)
                for sub in sub_list:
                    if sub[3] == "편성목":
                        continue
                    bullet = bool(sub[3] and sub[3].strip() and sub[3].strip()[0] in "○ㅇ■▣-")
                    next_is_label = _next_is_label_row(next_in_block)
                    if bullet and not sub[4] and next_in_block and not next_is_label:
                        amt = take_last_amount_from_line(next_in_block)
                        if amt:
                            sub[4] = amt
                            used_as_amount_only.add(fi + 1)
                    aoa.append(sub)
                    if sub[3] and not sub[4]:
                        last_row_with_no_amount = sub

        if used_next_export_line:
            j += 1
        j += 1

    return aoa
