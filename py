pension.pdf, health.pdf, emp_worker.pdf, ia_worker.pdf, emp_labor.pdf, ia_labor.pdf



!pip -q install pdfplumber openpyxl




import io
import re
from pathlib import Path
from collections import defaultdict

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from google.colab import files


# =========================
#  상단 문구(요청 반영)
# =========================
print("사회보험료 개별 고지서 엑셀 시트 변환 프로그램")
print("일 좀 덜 힘들게 해보자. 졸려죽겠네…")
print("-" * 60)


# =========================
#  설정/상수
# =========================
BAD_WORDS = {
    "건강","요양","총계","장기","국민","사업","직장","구분","고지","산출",
    "부과","출력","페이지","관리","번호","보험","료","정산","환급"
}

# 업로드 파일명 규칙(Colab에서 구분용)
# ✅ 업로드 시 파일명을 아래처럼 맞추면 자동 인식됨
# - pension.pdf     : 근로자 국민연금
# - health.pdf      : 근로자 건강보험(요양 포함)
# - emp_worker.pdf  : 근로자 고용보험
# - ia_worker.pdf   : 근로자 산재보험
# - emp_labor.pdf   : 노무제공자(방과후) 고용보험
# - ia_labor.pdf    : 노무제공자(방과후) 산재보험
EXPECTED = {
    "pension": "pension.pdf",
    "health": "health.pdf",
    "emp_worker": "emp_worker.pdf",
    "ia_worker": "ia_worker.pdf",
    "emp_labor": "emp_labor.pdf",
    "ia_labor": "ia_labor.pdf",
}


# =========================
#  공용 유틸
# =========================
def extract_text(pdf_path: Path) -> str:
    parts = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for p in pdf.pages:
            parts.append(p.extract_text() or "")
    return "\n".join(parts)

def money_to_int(s: str) -> int:
    return int(s.replace(",", "").strip())


# =========================
#  파서들
# =========================
def parse_pension(pdf_path: Path, master: dict):
    text = extract_text(pdf_path)
    pat = re.compile(
        r"^\s*(\d+)\s+\d+\s+\d{6}-\d\*+\s+([가-힣]+)\s+\d+\s+\d{4}\.\d{2}\s+~\s+\d{4}\.\d{2}\s+([\d,]+)\s*$",
        re.M
    )
    for m in pat.finditer(text):
        name = m.group(2).strip()
        total = money_to_int(m.group(3))
        emp = total // 2
        org = total - emp
        rec = master[name]
        rec["성명"] = name
        rec["국민연금_개인"] = emp
        rec["국민연금_기관"] = org

def parse_ia_worker(pdf_path: Path, master: dict):
    text = extract_text(pdf_path)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    for ln in lines:
        m = re.match(r"^\d+\s+일반\s+([가-힣]{2,4})\s+\d{2}-\d{2}-\d{2}.*\s(-?\d[\d,]*)\s*$", ln)
        if m:
            name = m.group(1).strip()
            amt = money_to_int(m.group(2))
            rec = master[name]
            rec["성명"] = name
            rec["산재_기관"] = amt

def parse_emp_worker(pdf_path: Path, master: dict):
    text = extract_text(pdf_path)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    current = None
    for ln in lines:
        m = re.match(r"^(\d+)\s+일반\s+([가-힣]{2,4})\s+\d{2}-\d{2}-\d{2}", ln)
        if m:
            current = m.group(2).strip()
            rec = master[current]
            rec["성명"] = current
            continue

        if not current:
            continue

        nums = re.findall(r"-?\d[\d,]*", ln)
        if not nums:
            continue

        if "근로자실업급여보험료" in ln:
            master[current]["고용_개인실업"] = money_to_int(nums[-1])
        elif "사업주실업급여보험료" in ln:
            master[current]["고용_기관실업"] = money_to_int(nums[-1])
        elif "사업주고안직능보험료" in ln:
            master[current]["고용_기관고안직능"] = money_to_int(nums[-1])

def parse_health(pdf_path: Path, master: dict):
    text = extract_text(pdf_path)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    for ln in lines:
        nm = re.search(r"\s([가-힣]{2,4})\s", " " + ln + " ")
        if not nm:
            continue
        name = nm.group(1).strip()
        if name in BAD_WORDS:
            continue

        is_health = bool(re.search(r"건\s*강|건강", ln))
        is_care = bool(re.search(r"요\s*양|요양", ln))
        if not (is_health or is_care):
            continue

        nums = re.findall(r"-?\d[\d,]*", ln)
        if not nums:
            continue
        amt = money_to_int(nums[-1])

        rec = master[name]
        rec["성명"] = name

        # MVP 로직(원본 그대로 유지)
        if is_health and not is_care:
            rec["건강_개인"] = amt
            rec["건강_기관"] = amt
        elif is_care and not is_health:
            rec["요양_개인"] = amt
            rec["요양_기관"] = amt

def parse_emp_labor(pdf_path: Path, master: dict):
    text = extract_text(pdf_path)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    tmp_emp = None
    tmp_org = None

    for ln in lines:
        if "근로자실업급여보험료" in ln:
            nums = re.findall(r"-?\d[\d,]*", ln)
            if nums:
                tmp_emp = money_to_int(nums[-1])
        elif "사업주실업급여보험료" in ln:
            nums = re.findall(r"-?\d[\d,]*", ln)
            if nums:
                tmp_org = money_to_int(nums[-1])

        m = re.match(r"^\d+\s+특수\(\)\s+([가-힣]{2,4})\s+\d{2}-\d{2}-\d{2}", ln)
        if m:
            name = m.group(1).strip()
            rec = master[name]
            rec["성명"] = name
            rec["노무고용_개인"] = tmp_emp or 0
            rec["노무고용_기관"] = tmp_org or 0
            tmp_emp = None
            tmp_org = None

def parse_ia_labor(pdf_path: Path, master: dict):
    text = extract_text(pdf_path)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    for ln in lines:
        m = re.match(r"^\d+\s+제공자\(\s*([가-힣]{2,4})\s+\d{2}-\d{2}-\d{2}.*\s([\d,]+)\s*$", ln)
        if m:
            name = m.group(1).strip()
            total = money_to_int(m.group(2))
            emp = total // 2
            org = total - emp
            rec = master[name]
            rec["성명"] = name
            rec["노무산재_개인"] = emp
            rec["노무산재_기관"] = org


# =========================
#  분류/합계/기본값
# =========================
def infer_class(rec: dict) -> str:
    labor_sum = (
        rec.get("노무고용_개인",0)+rec.get("노무고용_기관",0)+
        rec.get("노무산재_개인",0)+rec.get("노무산재_기관",0)
    )
    if labor_sum > 0:
        return "노무제공자"

    four_sum = (
        rec.get("국민연금_개인",0)+rec.get("국민연금_기관",0)+
        rec.get("건강_개인",0)+rec.get("건강_기관",0)+
        rec.get("요양_개인",0)+rec.get("요양_기관",0)+
        rec.get("고용_개인실업",0)+rec.get("고용_기관실업",0)+rec.get("고용_기관고안직능",0)
    )
    if four_sum == 0 and rec.get("산재_기관",0) > 0:
        return "공무직"
    return "근로자"

def ensure_defaults(master: dict):
    keys = [
        "국민연금_개인","국민연금_기관",
        "건강_개인","건강_기관",
        "요양_개인","요양_기관",
        "고용_개인실업","고용_기관실업","고용_기관고안직능",
        "산재_기관",
        "노무고용_개인","노무고용_기관",
        "노무산재_개인","노무산재_기관",
    ]
    for name, rec in master.items():
        rec.setdefault("성명", name)
        for k in keys:
            rec.setdefault(k, 0)

def compute_sums(rec: dict):
    if rec["구분"] in ("근로자","공무직"):
        personal = rec["국민연금_개인"] + rec["건강_개인"] + rec["요양_개인"] + rec["고용_개인실업"]
        org = rec["국민연금_기관"] + rec["건강_기관"] + rec["요양_기관"] + rec["고용_기관실업"] + rec["고용_기관고안직능"] + rec["산재_기관"]
    else:
        personal = rec["노무고용_개인"] + rec["노무산재_개인"]
        org = rec["노무고용_기관"] + rec["노무산재_기관"]

    rec["개인부담금"] = personal
    rec["기관부담금"] = org
    rec["총액"] = personal + org


# =========================
#  엑셀 생성
# =========================
def style_worksheet(ws, headers, rows):
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="F2F2F2")
    header_font = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    ws.freeze_panes = "A2"

    for j, h in enumerate(headers, start=1):
        c = ws.cell(1, j, h)
        c.fill = header_fill
        c.font = header_font
        c.border = border
        c.alignment = center

    for i, row in enumerate(rows, start=2):
        for j, val in enumerate(row, start=1):
            c = ws.cell(i, j, val)
            c.border = border
            if isinstance(val, int):
                c.number_format = "#,##0"
                c.alignment = right
            else:
                c.alignment = left if j == 1 else center

    for j, h in enumerate(headers, start=1):
        maxlen = len(str(h))
        for r in rows[:60]:
            maxlen = max(maxlen, len(str(r[j-1])))
        ws.column_dimensions[get_column_letter(j)].width = min(max(12, maxlen + 2), 34)

def build_excel(records) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    workers = [r for r in records if r["구분"] == "근로자"]
    public = [r for r in records if r["구분"] == "공무직"]
    labor  = [r for r in records if r["구분"] == "노무제공자"]

    ws = wb.create_sheet("요약", 0)
    ws["A1"] = "사회보험료 산출(Colab)"
    ws["A1"].font = Font(bold=True, size=14)
    ws.append(["구분","인원","개인부담금","기관부담금","총액"])
    for cell in ws[2]:
        cell.font = Font(bold=True)

    def sums(arr):
        return (len(arr),
                sum(x["개인부담금"] for x in arr),
                sum(x["기관부담금"] for x in arr),
                sum(x["총액"] for x in arr))

    ws.append(["근로자(교육공무직 제외)", *sums(workers)])
    ws.append(["교육공무직 산재", *sums(public)])
    ws.append(["노무제공자(방과후)", *sums(labor)])

    for r in range(3,6):
        for c in ("C","D","E"):
            ws[f"{c}{r}"].number_format = "#,##0"

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 16

    ws1 = wb.create_sheet("근로자(교육공무직 제외)")
    h1 = ["성명","국민연금(개인)","국민연금(기관)","건강(개인)","건강(기관)","요양(개인)","요양(기관)",
          "고용(개인-실업)","고용(기관-실업)","고용(기관-고안직능)","산재(기관)","개인부담금","기관부담금","총액"]
    r1 = [[x["성명"], x["국민연금_개인"], x["국민연금_기관"], x["건강_개인"], x["건강_기관"],
           x["요양_개인"], x["요양_기관"], x["고용_개인실업"], x["고용_기관실업"], x["고용_기관고안직능"],
           x["산재_기관"], x["개인부담금"], x["기관부담금"], x["총액"]] for x in workers]
    style_worksheet(ws1, h1, r1)

    ws2 = wb.create_sheet("교육공무직 산재")
    h2 = ["성명","산재보험료(기관)","기관부담금","총액"]
    r2 = [[x["성명"], x["산재_기관"], x["기관부담금"], x["총액"]] for x in public]
    style_worksheet(ws2, h2, r2)

    ws3 = wb.create_sheet("노무제공자(방과후)")
    h3 = ["성명","노무고용(개인)","노무고용(기관)","노무산재(개인)","노무산재(기관)","개인부담금","기관부담금","총액"]
    r3 = [[x["성명"], x["노무고용_개인"], x["노무고용_기관"], x["노무산재_개인"], x["노무산재_기관"],
           x["개인부담금"], x["기관부담금"], x["총액"]] for x in labor]
    style_worksheet(ws3, h3, r3)

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


# =========================
#  실행: 업로드 → 변환 → 다운로드
# =========================
print("1) PDF 업로드를 시작합니다.")
print("   권장: 파일명을 아래처럼 맞춰서 올리면 자동 인식됩니다.")
for k, fn in EXPECTED.items():
    print(f"   - {fn}")

uploaded = files.upload()  # 모바일에서도 버튼으로 업로드 됨

# 업로드된 파일들을 Path로 매핑
pdf_paths = {k: None for k in EXPECTED.keys()}
for key, expected_name in EXPECTED.items():
    if expected_name in uploaded:
        pdf_paths[key] = Path(expected_name)

# 하나도 매칭이 안 되면: 업로드된 pdf를 보여주고 사용자에게 파일명 변경을 유도
if not any(pdf_paths.values()):
    pdfs = [name for name in uploaded.keys() if name.lower().endswith(".pdf")]
    print("\n⚠️ 업로드된 PDF 목록:", pdfs)
    raise ValueError(
        "업로드된 파일명이 규칙과 매칭되지 않습니다.\n"
        "PDF 파일명을 pension.pdf / health.pdf / emp_worker.pdf / ia_worker.pdf / emp_labor.pdf / ia_labor.pdf 로 바꿔서 다시 업로드해 주세요."
    )

print("\n2) 변환을 시작합니다...")

master = defaultdict(dict)

if pdf_paths["pension"]:
    parse_pension(pdf_paths["pension"], master)
if pdf_paths["health"]:
    parse_health(pdf_paths["health"], master)
if pdf_paths["emp_worker"]:
    parse_emp_worker(pdf_paths["emp_worker"], master)
if pdf_paths["ia_worker"]:
    parse_ia_worker(pdf_paths["ia_worker"], master)
if pdf_paths["emp_labor"]:
    parse_emp_labor(pdf_paths["emp_labor"], master)
if pdf_paths["ia_labor"]:
    parse_ia_labor(pdf_paths["ia_labor"], master)

ensure_defaults(master)

records = []
for name in sorted(master.keys()):
    rec = master[name]
    rec["구분"] = infer_class(rec)
    compute_sums(rec)
    records.append(rec)

xlsx_bytes = build_excel(records)

out_name = "사회보험료_산출.xlsx"
Path(out_name).write_bytes(xlsx_bytes)

print("3) 엑셀 생성 완료:", out_name)
files.download(out_name)