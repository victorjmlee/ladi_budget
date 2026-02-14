#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
웹 UI 제공 + 엑셀 추출 API (Python 파이프라인).
실행: python app.py  →  http://localhost:3580
"""
import io
import os
import tempfile
from pathlib import Path

from flask import Flask, request, send_file, send_from_directory

# 프로젝트 루트(kolo)에서 import
from pdf_extract import extract_text_by_page, extract_by_codes, extract_by_keywords
from excel_build import build_excel_aoa

app = Flask(__name__, static_folder=None)
ROOT = Path(__file__).resolve().parent
PORT = int(os.environ.get("PORT", 3580))


def parse_keywords(raw: str):
    return [s.strip() for s in raw.replace(",", "\n").splitlines() if s.strip()]


@app.route("/")
def index():
    return send_from_directory(ROOT, "index.html")


@app.route("/<path:path>")
def static_file(path):
    if path == "index.html":
        return send_from_directory(ROOT, path)
    if Path(ROOT / path).is_file():
        return send_from_directory(ROOT, path)
    return "", 404


@app.route("/api/extract-excel", methods=["POST"])
def extract_excel():
    if "pdf" not in request.files:
        return {"error": "PDF 파일이 없습니다."}, 400
    file = request.files["pdf"]
    if not file or not file.filename or not file.filename.lower().endswith(".pdf"):
        return {"error": "PDF 파일을 선택하세요."}, 400
    raw_codes = (request.form.get("codes") or "").strip()
    raw_keywords = (request.form.get("keywords") or "").strip()
    codes = [s for s in parse_keywords(raw_codes) if len(s) == 3 and s.isdigit()]
    kw_list = [s for s in parse_keywords(raw_keywords) if s]
    # 하위호환: codes 필드 없이 keywords에 3자리 숫자만 온 경우 코드로 처리
    if not codes and kw_list and all(len(k) == 3 and k.isdigit() for k in kw_list):
        codes = kw_list
        kw_list = []
    if not codes and not kw_list:
        return {"error": "코드(201, 207) 또는 키워드를 입력하세요."}, 400

    try:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            file.save(tmp.name)
            tmp_path = tmp.name
        try:
            page_texts = extract_text_by_page(tmp_path)
            code_results = extract_by_codes(page_texts, codes) if codes else []
            kw_results = extract_by_keywords(page_texts, kw_list) if kw_list else []
            if not code_results and not kw_results:
                return {"error": "매칭되는 줄이 없습니다."}, 400
            code_aoa = build_excel_aoa(code_results, page_texts) if code_results else None
            kw_aoa = build_excel_aoa(kw_results, page_texts) if kw_results else None
        finally:
            try:
                os.unlink(tmp_path)
            except Exception:
                pass

        import openpyxl
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter
        from parser import EXCEL_HEADER
        from pdf_extract import _search_norm

        def _match_kws(text, keywords):
            n = _search_norm(text)
            return [kw for kw in keywords if _search_norm(kw) in n]

        combined = [EXCEL_HEADER + ["키워드검수"]]

        if code_aoa:
            # 코드 결과 = 메인. 각 행에 키워드 매칭 여부 표시
            code_text_norms = set()
            for row in code_aoa[1:]:
                matched = _match_kws(row[3], kw_list) if row[3] and kw_list else []
                combined.append(row + [", ".join(matched)])
                if row[3] and row[3].strip() != "편성목":
                    code_text_norms.add(_search_norm(row[3]))

            # 키워드로만 잡힌 항목 (코드로 안 잡힌 것) → 놓친 항목
            if kw_results:
                seen = set()
                kw_only = []
                for line, pi, li in kw_results:
                    n = _search_norm(line)
                    if n in seen:
                        continue
                    in_code = any(n in cn or cn in n for cn in code_text_norms)
                    if not in_code:
                        seen.add(n)
                        kw_only.append((line, pi, li))
                if kw_only:
                    kw_only_aoa = build_excel_aoa(kw_only, page_texts)
                    combined.append([""] * 8)
                    combined.append(["", "", "",
                                     "── 코드 미포함 (키워드로만 검색됨) ──",
                                     "", "", "", ""])
                    for row in kw_only_aoa[1:]:
                        matched = _match_kws(row[3], kw_list) if row[3] else []
                        combined.append(row + [", ".join(matched)])

        elif kw_aoa:
            # 키워드만 검색 (코드 없음)
            for row in kw_aoa[1:]:
                matched = _match_kws(row[3], kw_list) if row[3] else []
                combined.append(row + [", ".join(matched)])

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "예산추출"
        header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        verify_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        miss_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

        in_miss = False
        for r, row in enumerate(combined, 1):
            is_sep = len(row) > 3 and "코드 미포함" in (row[3] or "")
            if is_sep:
                in_miss = True
            for c, val in enumerate(row, 1):
                cell = ws.cell(row=r, column=c, value=val or "")
                if r == 1:
                    cell.font = Font(bold=True)
                    cell.fill = header_fill
                elif is_sep:
                    cell.font = Font(bold=True)
                elif in_miss:
                    cell.fill = miss_fill
                elif c == 8 and val:
                    cell.fill = verify_fill

        ws.auto_filter.ref = f"A1:H{len(combined)}"

        for col_idx in range(1, 9):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for row_idx in range(1, len(combined) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                val = cell.value
                if val is not None:
                    s = str(val)
                    length = sum(2 if "\uAC00" <= ch <= "\uD7A3" or ord(ch) > 127 else 1 for ch in s)
                    max_len = max(max_len, length)
            ws.column_dimensions[col_letter].width = min(max(max_len / 2 + 2, 8), 55)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        base_name = Path(file.filename).stem or "검색결과"
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=base_name + ".xlsx",
        )
    except Exception as e:
        return {"error": str(e)}, 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT, debug=False)
