# -*- coding: utf-8 -*-
"""
PDF에서 페이지별 텍스트 추출, 코드/키워드 필터.
"""
import re
from pathlib import Path
from typing import List, Tuple

try:
    import fitz  # pymupdf
except ImportError:
    fitz = None

from parser import (
    normalize_extracted_text,
    is_3digit_start_line,
    is_new_section_title,
)

def _is_page_header_simple(line: str) -> bool:
    """페이지 헤더 간단 체크 (블록 수집 중)."""
    t = (line or "").strip()
    if not t:
        return True
    # 부서/정책/단위 헤더
    if re.match(r"^부서\s*[:：]", t) or re.match(r"^정책\s*[:：]", t) or re.match(r"^단위\s*[:：]", t):
        return True
    if re.match(r"^\(단위\s*[:：]?\s*천원\)", t) or "(단위:천원)" in t:
        return True
    # 테이블 헤더
    if "부서ㆍ정책ㆍ단위" in t and "편성목" in t:
        return True
    if "세부사업" in t and "편성목" in t and "예산액" in t:
        return True
    if "예산액" in t and "전년도" in t and "비교증감" in t:
        return True
    if t in ("예산액", "전년도", "비교증감"):
        return True
    return False


def _page_text_from_blocks(page) -> str:
    """블록 단위 추출, (y0,x0) 정렬로 읽기 순서 유지. 줄마다 정규화."""
    blocks = page.get_text("blocks")
    blocks = sorted(blocks, key=lambda b: (round(b[1], 1), round(b[0], 1)) if len(b) >= 4 else (0, 0))
    lines = []
    for b in blocks:
        if len(b) >= 5 and b[4]:
            for line in b[4].strip().splitlines():
                line = line.strip()
                if line:
                    lines.append(normalize_extracted_text(line))
    return "\n".join(lines)


def extract_text_by_page(pdf_path: str) -> List[str]:
    """PDF 경로에서 페이지별로 텍스트 추출. 블록→줄 단위 유지 후 줄마다 정규화."""
    if not fitz:
        raise RuntimeError("pymupdf 필요: pip install pymupdf")
    path = Path(pdf_path)
    if not path.exists():
        raise FileNotFoundError(str(path))
    doc = fitz.open(str(path))
    pages = []
    try:
        for i in range(len(doc)):
            page = doc.load_page(i)
            text = _page_text_from_blocks(page)
            if not text.strip():
                text = page.get_text("text")
                text = normalize_extracted_text(text)
            pages.append(text)
    finally:
        doc.close()
    return pages


def get_page_lines(page_text: str) -> List[str]:
    return [s.strip() for s in (page_text or "").splitlines() if s.strip()]


def get_line_index_in_page(line: str, page_lines: List[str]) -> int:
    """페이지 라인 배열에서 해당 줄의 인덱스. 정규화 비교."""
    def norm(s: str) -> str:
        return (s or "").replace(" ", "").replace("ㅇ", "○")
    key = norm(line)
    for i, pl in enumerate(page_lines):
        if norm(pl) == key or norm(pl) == (line or "").replace(" ", ""):
            return i
    if len(key) >= 4:
        for i, pl in enumerate(page_lines):
            if key in norm(pl) or norm(pl) in key:
                return i
    return -1


def get_block_start(page_lines: List[str], line_index: int) -> int:
    """해당 줄이 속한 블록의 시작(직전 3자리 코드) 인덱스."""
    if line_index < 0 or line_index >= len(page_lines):
        return -1
    for i in range(line_index, -1, -1):
        if is_3digit_start_line(page_lines[i]):
            return i
    return -1


def get_next_lines_until_next_block(
    page_texts: List[str],
    page_index: int,
    line_index: int,
) -> List[str]:
    """
    현재 페이지 line_index 다음 줄부터 수집.
    종료 조건:
    - 다음 3자리 코드 전까지
    - 또는 '['로 시작하는 새 섹션 전까지
    - 또는 새 단위명(문화행사지원 등) 전까지
    
    "-" 로 시작하는 줄은 여러 개 포함 가능 (원고료, 번역및조사비 등)
    """
    out: List[str] = []
    total = len(page_texts)
    from_idx = line_index + 1
    p = page_index
    while p < total:
        lines = get_page_lines(page_texts[p])
        start = from_idx if p == page_index else 0
        for k in range(start, len(lines)):
            line = lines[k]
            # 다음 3자리 코드면 끝
            if is_3digit_start_line(line):
                return out
            # [기금], [보조] 등 새 섹션이면 그 전에서 끊음
            if line.strip().startswith('['):
                return out
            # 새 단위명이면 그 전에서 끊음 (포함 안 함)
            if is_new_section_title(line):
                return out
            # 페이지 헤더면 스킵하고 계속 (페이지 넘어가도 블록 유지)
            if _is_page_header_simple(line):
                continue
            out.append(line)
        p += 1
        from_idx = 0
    return out


def _first_column_is_code(line: str, code: str) -> bool:
    """첫 컬럼(2칸 이상 공백 구분)이 예산 코드인지 확인. 금액란의 207(2,070, 207,000 등)은 제외."""
    cols = re.split(r"\s{2,}", (line or "").strip())
    cols = [c.replace(" ", "").strip() for c in cols if c.strip()]
    if not cols or len(cols[0]) < 3:
        return False
    first = cols[0]
    if first[:3] != code or not first[:3].isdigit():
        return False
    # 207연구개발비, 207 → O / 207,000, 2,070 → X (코드 뒤가 금액이면 제외)
    if first == code:
        return True
    rest = first[3:]
    return bool(re.search(r"[가-힣]", rest))


def extract_by_codes(
    page_texts: List[str], codes: List[str]
) -> List[Tuple[str, int, int]]:
    """코드(201, 207 등)로 매칭되는 줄만 (텍스트, 페이지인덱스, 줄인덱스) 리스트로 반환.
    첫 컬럼이 해당 코드일 때만 포함(금액란에 207 들어간 줄은 제외).
    여러 코드 입력 시 코드별로 묶어서 반환: 201 전부 → 207 전부 (페이지 순 유지)."""
    codes_ordered = [c.strip() for c in codes if c.strip()]
    result: List[Tuple[str, int, int]] = []
    
    for code in codes_ordered:
        for page_index, text in enumerate(page_texts):
            page_lines = get_page_lines(text)
            for line_index, line in enumerate(page_lines):
                if not _first_column_is_code(line, code):
                    continue
                result.append((line, page_index, line_index))
    return result


def extract_by_keywords(
    page_texts: List[str], keywords: List[str]
) -> List[Tuple[str, int, int]]:
    """키워드가 포함된 줄 (및 그 주변) 수집. (텍스트, 페이지인덱스, 줄인덱스) 반환."""
    result: List[Tuple[str, int, int]] = []
    norm = lambda s: (s or "").replace(" ", "").replace("○", "○").replace("ㅇ", "○")
    kw_norms = [norm(k) for k in keywords if k.strip()]
    for page_index, text in enumerate(page_texts):
        page_lines = get_page_lines(text)
        for line_index, line in enumerate(page_lines):
            n = norm(line)
            for kn in kw_norms:
                if kn in n or n in kn:
                    result.append((line, page_index, line_index))
                    break
    return result
