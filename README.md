# kolo – PDF 키워드 추출 (견적용)

PDF에서 키워드에 해당하는 내용만 뽑아 견적 형태로 보거나 다운로드합니다.

## 웹 앱

**Python으로 실행 (추천)** – 엑셀 다운로드 시 서버에서 Python 추출(블록 누락 없음):

```bash
cd ~/kolo
pip install -r requirements.txt   # 최초 1회
python app.py
```

**Node로 실행** – 엑셀은 브라우저에서 생성:

```bash
cd ~/kolo
npm run web
```

둘 다 **http://localhost:3580** 에서 동일한 화면으로 사용합니다.

1. **PDF 업로드** → 왼쪽에 전체 내용 표시
2. **키워드/코드 입력** (쉼표 또는 줄바꿈, 예: `201, 207` 또는 `물가`) → 오른쪽에 추출 결과 표시
3. **엑셀 다운로드** → `python app.py` 일 때는 서버(Python) 추출, `npm run web` 일 때는 브라우저 추출

## CLI

```bash
npm install   # 최초 1회
npm run extract
```

PDF 경로 지정: `node extract.mjs "파일이름.pdf"`

## 키워드 변경

- **웹**: 텍스트 칸에 원하는 키워드를 입력하면 됩니다.
- **CLI**: `extract.mjs` 상단의 `KEYWORDS` 배열을 수정하세요.

---

## Python으로 엑셀 추출 (권장)

규칙을 코드로 관리하고, **페이지를 넘어가는 블록**(▣운영수당·물가대책·착한가격 등)까지 한 번에 수집할 수 있습니다.

```bash
cd ~/kolo
pip install -r requirements.txt
```

**코드로 추출 (3자리 예산 코드):**

```bash
python run.py 세출예산사업명세서\(일자리경제과\).pdf --codes 201,207,401 --out result.xlsx
```

**키워드로 추출:**

```bash
python run.py 세출예산사업명세서.pdf --keywords 물가,착한가격 --out result.xlsx
```

- 출력 엑셀 컬럼: 부서, 정책, 단위, 세부사업, 예산액, 전년도예산액, 비교증감
- 파싱 규칙은 `parser.py`, 블록 수집은 `pdf_extract.py`, 엑셀 행 구성은 `excel_build.py`에 있어 수정·테스트가 쉽습니다.
