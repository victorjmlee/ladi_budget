# -*- coding: utf-8 -*-
"""
세출예산사업명세서 PDF 추출 텍스트 파싱.
열 구분: 스페이스 2개 이상. 한글 구간 공백 제거, 숫자/금액 패턴 유지.
"""
import re
from typing import List, Optional

EXCEL_HEADER = ["부서", "정책", "단위", "세부사업", "예산액", "전년도예산액", "비교증감"]

AMOUNT_RE = re.compile(
    r"^\d{1,3}(?:,\d{3})+$|^△\d{1,3}(?:,\d{3})*$|^△?\d{1,6}$"
)
ONLY_NUMBER_RE = re.compile(
    r"^\d{1,3}(?:,\d{3})+$|^\d{3,6}$"
)
BULLET_START_RE = re.compile(r"^[▣■○ㅇ]")
FORMULA_OR_AMOUNT_LINE_RE = re.compile(
    r"^[가-힣▣■○ㅇ]"
)  # 반대: 이걸로 시작 안 하면 식+금액 줄 후보
TWO_DIGIT_START_RE = re.compile(r"^\d{2}\s")
THREE_DIGIT_START_RE = re.compile(r"^\d{3}$|^\d{3}[가-힣]")
NEW_SECTION_TITLE_RE = re.compile(
    r"^\d{1,3}(?:,\d{3})+$|^△\d"
)


def _norm(s: str) -> str:
    return (s or "").replace(" ", "").replace("○", "○").replace("ㅇ", "○")


def collapse_spaces_in_korean(text: str) -> str:
    if not text:
        return text
    return re.sub(r"([가-힣](?:\s+[가-힣])*)", lambda m: m.group(1).replace(" ", ""), text)


def normalize_extracted_text(text: str) -> str:
    """한글 구간 공백 제거, 2칸 이상은 열 구분으로 유지. 끝의 ' 금액'은 한 칸이라도 유지."""
    if not text:
        return text
    s = collapse_spaces_in_korean(text)
    # 끝의 ' 숫자/금액'은 한 칸이라도 열 구분이 되도록 보호
    trailing_amount = re.search(r"\s+(\d{1,3}(?:,\d{3})*|△\d[\d,]*)\s*$", s)
    if trailing_amount:
        s = s[: trailing_amount.start()] + "\u200C\u200C" + trailing_amount.group(1)
    s = re.sub(r"  +", "\u200B\u200B", s)
    s = s.replace(" ", "")
    s = s.replace("\u200B\u200B", "  ")
    s = re.sub(r",\s+", ",", s)
    s = s.replace("\u200C\u200C", "  ")
    return s


def is_amount(s: str) -> bool:
    t = (s or "").strip()
    if not t:
        return False
    if re.match(r"^\d{1,3}(?:,\d{3})+$", t):
        return True
    if re.match(r"^△\d{1,3}(?:,\d{3})*$", t):
        return True
    if re.match(r"^△?\d{1,6}$", t):
        return True
    return False


def format_sebu_label(s: str) -> str:
    m = re.match(r"^(\d{2,3})([가-힣].*)$", (s or "").strip())
    return f"{m.group(1)} {m.group(2)}" if m else (s or "").strip()


def _split_cols(line: str) -> List[str]:
    """2칸 이상 공백으로 열 구분. 한 칸 공백이어도 '라벨 + 금액'이면 열 구분."""
    line = (line or "").strip()
    parts = re.split(r"\s{2,}", line)
    parts = [p.replace(" ", "").strip() for p in parts if p.strip()]
    if len(parts) == 1 and " " in line:
        last_space = line.rfind(" ")
        if last_space > 0:
            after = line[last_space + 1 :].replace(" ", "").strip()
            if after and is_amount(after):
                before = line[:last_space].replace(" ", "").strip()
                if before:
                    return [before, after]
    return parts


def take_last_amount_from_line(line: str) -> Optional[str]:
    """2칸 이상 공백 뒤 마지막 열만 예산액. 식(100,000원*11명*2회) 안의 숫자는 제외."""
    if not line:
        return None
    parts = _split_cols(line)
    if len(parts) >= 2:
        last = parts[-1]
        if re.match(r"^\d{1,3}(?:,\d{3})*$|^△\d{1,3}(?:,\d{3})*$", last):
            return last
    collapsed = re.sub(r"\s+", "", line)
    m = re.search(r"(\d{1,3}(?:,\d{3})*)\s*$", collapsed)
    if m:
        return m.group(1)
    return None


def take_first_amount_from_line(line: str) -> Optional[str]:
    if not line:
        return None
    collapsed = re.sub(r"\s+", "", line)
    m = re.search(r"\d{1,3}(?:,\d{3})+", collapsed) or re.search(r"\d{1,3}(?:,\d{3})*", collapsed)
    return m.group(0) if m else None


def is_3digit_start_line(line: str) -> bool:
    """세부사업 코드(201, 202 등)만 블록 시작. 금액(395), 식(100,000원*…) 제외."""
    cols = _split_cols(line.strip())
    if not cols:
        return False
    c0 = cols[0]
    if "," in c0:
        return False
    if re.match(r"^\d{3}$", c0):
        return bool(re.search(r"[가-힣]", line))
    if re.match(r"^\d{3}", c0) and re.search(r"[가-힣]", c0):
        return True
    return False


def is_new_section_title(line: str) -> bool:
    """새 단위/사업 구간 제목(예: 전통시장 활성화사업 추진, 문화행사 지원)이면 True. 블록 수집 여기서 끊음."""
    t = (line or "").strip()
    if not t or re.match(r"^\d{2,3}\s", t) or re.match(r"^[▣■○ㅇ\-]", t):
        return False
    # 전통시장 활성화사업 등 한글로만 된 사업명(금액이 다음 줄에 있어도) → 새 섹션
    if re.search(r"전통시장\s*활성화", t) and re.search(r"[가-힣]", t):
        return True
    # 한글로만 된 단위명(문화행사 지원, 영상산업 관리 등) - 금액이 다음 줄이어도 새 섹션으로 처리
    t_no_space = t.replace(" ", "")
    if re.match(r"^[가-힣]+$", t_no_space) and len(t_no_space) >= 3:
        if re.search(r"(관리|지원|추진|운영|사업|조성|육성|활성화|개발|보급|확대)$", t_no_space):
            return True
    cols = _split_cols(t)
    if len(cols) < 2:
        return False
    first = cols[0]
    if not re.search(r"[가-힣]", first) or re.match(r"^\d", first):
        return False
    last = cols[-1]
    return bool(re.match(r"^\d{1,3}(?:,\d{3})+$", last)) or bool(re.match(r"^△\d", last))


def _empty_row() -> List[str]:
    return [""] * 7


def _pyeon_header_row() -> List[str]:
    r = _empty_row()
    r[3] = "편성목"
    return r


def _is_sebu_code(col0: str) -> bool:
    return bool(re.match(r"^\d{3}$", col0)) or (
        re.match(r"^\d{3}", col0) and re.search(r"[가-힣]", col0)
    )


def extract_pdf_meta(full_text: str) -> dict:
    """페이지 텍스트에서 부서/정책/단위 추출."""
    meta = {"부서": "", "정책": "", "단위": ""}
    def _norm_meta(s: str) -> str:
        return re.sub(r"\s+", " ", (s or "")).strip()
    t = (full_text or "").strip()
    m1 = re.search(r"부서\s*[:：]\s*(.+?)정책\s*[:：]", t, re.DOTALL)
    m2 = re.search(r"정책\s*[:：]\s*(.+?)단위\s*[:：]", t, re.DOTALL)
    m3 = re.search(r"단위\s*[:：]\s*([^\n]*?)(?=\n|$)", t)
    if m1:
        meta["부서"] = _norm_meta(m1.group(1))
    if m2:
        meta["정책"] = _norm_meta(m2.group(1))
    if m3:
        meta["단위"] = _norm_meta(m3.group(1))
    if not (meta["부서"] or meta["정책"] or meta["단위"]):
        for pat, key in [(r"부\s*서", "부서"), (r"정\s*책", "정책"), (r"단\s*위", "단위")]:
            a = re.search(pat + r"\s*[:：]\s*([^\n]+)", t)
            if a:
                meta[key] = _norm_meta(a.group(1))
    if not (meta["부서"] or meta["정책"] or meta["단위"]):
        for key in ["부서", "정책", "단위"]:
            b = re.search(rf"{key}\s*[:：]\s*([^\n]+)", t)
            if b:
                meta[key] = _norm_meta(b.group(1))
    return meta


def apply_default_meta(meta: Optional[dict], fallback: Optional[dict]) -> dict:
    out = {}
    for k in ("부서", "정책", "단위"):
        out[k] = (meta or {}).get(k) or (fallback or {}).get(k) or ""
    return out


def line_to_excel_rows(line: str) -> List[List[str]]:
    """한 줄을 엑셀 행(들)로 변환. 7열: 부서,정책,단위,세부사업,예산액,전년도,비교증감."""
    line = (line or "").strip()
    cols = _split_cols(line)
    rows: List[List[str]] = []

    # 5열: 코드, 세부사업명, 예산액, 전년도, 비교증감
    if (
        len(cols) >= 5
        and re.match(r"^\d{2,3}$", cols[0])
        and re.search(r"[가-힣]", cols[1])
        and is_amount(cols[2])
        and is_amount(cols[3])
        and (is_amount(cols[4]) or cols[4].startswith("△"))
    ):
        row = _empty_row()
        row[3] = f"{cols[0]} {cols[1]}" if len(cols[0]) == 3 else cols[0] + cols[1]
        row[4], row[5], row[6] = cols[2], cols[3], cols[4]
        rows.append(row)
        if _is_sebu_code(cols[0]):
            rows.append(_pyeon_header_row())
        return rows

    # 4열: 코드+명, 예산, 전년도, 비교증감
    if (
        len(cols) >= 4
        and is_amount(cols[1])
        and is_amount(cols[2])
        and (is_amount(cols[3]) or cols[3].startswith("△"))
    ):
        row = _empty_row()
        row[3] = format_sebu_label(cols[0])
        row[4], row[5], row[6] = cols[1], cols[2], cols[3]
        rows.append(row)
        if _is_sebu_code(cols[0]):
            rows.append(_pyeon_header_row())
        return rows

    # ▣■○ㅇ + 금액
    if len(cols) >= 2 and re.match(r"^[▣■○ㅇ]", cols[0]):
        last = cols[-1]
        if is_amount(last):
            row = _empty_row()
            row[3] = " ".join(cols[:-1]).replace("  ", " ").strip()
            row[4] = last
            return [row]

    # 2열 일반
    if len(cols) == 2 and is_amount(cols[1]):
        row = _empty_row()
        row[3] = format_sebu_label(cols[0])
        row[4] = cols[1]
        return [row]

    # 3열
    if len(cols) == 3 and is_amount(cols[1]) and is_amount(cols[2]):
        row = _empty_row()
        row[3] = format_sebu_label(cols[0])
        row[4], row[5] = cols[1], cols[2]
        rows.append(row)
        if re.match(r"^\d{3}", cols[0]):
            rows.append(_pyeon_header_row())
        return rows

    if (
        len(cols) == 3
        and re.match(r"^\d{2,3}$", cols[0])
        and re.search(r"[가-힣]", cols[1])
        and is_amount(cols[2])
    ):
        row = _empty_row()
        row[3] = cols[0] + cols[1]
        row[4] = cols[2]
        return [row]

    # 폴백: 불릿(-, ■○▣ㅇ)만 있거나 라벨만
    line_norm = re.sub(r"\s+", " ", line).strip()
    only_bullet = bool(re.match(r"^[■○▣ㅇ\-]\s*.+", line_norm))
    row = _empty_row()
    if only_bullet:
        row[3] = re.sub(
            r"\s*△?\d{1,3}(?:,\d{3})*(?:\s*△?\d{1,3}(?:,\d{3})*)*\s*$",
            "",
            line_norm,
        ).strip()
    else:
        row[3] = line_norm
    return [row]
