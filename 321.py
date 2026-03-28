# 321.py
import altair as alt
import io
import os
import pickle
import hashlib
from dataclasses import dataclass
from datetime import datetime, timezone, timedelta

import pandas as pd
import streamlit as st
import yaml
import streamlit_authenticator as stauth

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.barcharts import VerticalBarChart


# ===================== 路徑/基本設定 =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.yaml")
STORE_PATH = os.path.join(BASE_DIR, "grades_store.pkl")
LOG_PATH = os.path.join(BASE_DIR, "query_log.csv")

TZ_TAIPEI = timezone(timedelta(hours=8))

DEFAULT_SUBJECT_ROW = 0
DEFAULT_EVAL_ROW = 1
DEFAULT_HEADER_ROW = 2


# ===================== 小工具 =====================
def now_taipei_str() -> str:
    return datetime.now(TZ_TAIPEI).strftime("%Y-%m-%d %H:%M:%S")


def sha256_hex(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()[:12]


def append_log(event: dict):
    df = pd.DataFrame([event])
    if os.path.exists(LOG_PATH):
        old = pd.read_csv(LOG_PATH, encoding="utf-8")
        out = pd.concat([old, df], ignore_index=True)
    else:
        out = df
    out.to_csv(LOG_PATH, index=False, encoding="utf-8")


def save_store(obj):
    with open(STORE_PATH, "wb") as f:
        pickle.dump(obj, f)


def load_store():
    if not os.path.exists(STORE_PATH):
        return None
    with open(STORE_PATH, "rb") as f:
        obj = pickle.load(f)

    # 向下相容舊版 store：可能只有 excel_bytes，沒有 meta
    if isinstance(obj, dict):
        if "meta" not in obj or not isinstance(obj.get("meta"), dict):
            obj["meta"] = {}
        meta = obj["meta"]
        meta.setdefault("version", "legacy")
        meta.setdefault("updated_at", datetime.now(TZ_TAIPEI).strftime("%Y-%m-%d %H:%M:%S"))
        meta.setdefault("title_text", "成績")
        meta.setdefault("sheet", 0)
        meta.setdefault("subject_row", DEFAULT_SUBJECT_ROW)
        meta.setdefault("eval_row", DEFAULT_EVAL_ROW)
        meta.setdefault("header_row", DEFAULT_HEADER_ROW)
        meta.setdefault("rows", 0)
    return obj


def seat_to_int_safe(seat: str) -> int:
    try:
        return int(float(seat))
    except Exception:
        return 9999


# ===================== 字型註冊（PDF中文） =====================
@st.cache_resource
def register_chinese_font():
    candidates = [
        "msjh.ttc", "msjh.ttf", "mingliu.ttc",
        "simsun.ttc", "kaiu.ttf", "NotoSansCJKtc-Regular.otf"
    ]
    win_fonts = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    for name in candidates:
        path = os.path.join(win_fonts, name)
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("CJK", path))
                return "CJK"
            except Exception:
                pass
    pdfmetrics.registerFont(UnicodeCIDFont("MSung-Light"))
    return "MSung-Light"


FONT = register_chinese_font()


# ===================== 資料清理 =====================
def clean_text(x) -> str:
    s = str(x) if x is not None else ""
    s = s.replace("\u3000", " ").strip()
    s = " ".join(s.split())
    return s


def seat_to_str(v) -> str:
    s = clean_text(v)
    if s == "" or s.lower() in ("nan", "none"):
        return ""
    try:
        f = float(s)
        return str(int(f)) if f.is_integer() else s
    except Exception:
        return s


def to_float_or_none(s: str):
    s = clean_text(s)
    if s in ("", "-", "—", "－", "缺考", "免試", "請假", "未交", "缺", "無"):
        return None
    try:
        return float(s)
    except Exception:
        return None


def is_hidden_score(s: str) -> bool:
    s = clean_text(s)
    hidden_values = {
        "", "nan", "none", "-", "—", "－",
        "缺考", "免試", "請假", "未交", "缺", "無"
    }
    return s.lower() in hidden_values or s in hidden_values


# ===================== 解析Excel =====================
def parse_all_scores_from_bytes(excel_bytes: bytes, sheet_name, subject_row, eval_row, header_row):
    df = pd.read_excel(io.BytesIO(excel_bytes), header=None, sheet_name=sheet_name)

    subjects = df.iloc[subject_row].fillna("").map(clean_text).tolist()
    evals = df.iloc[eval_row].fillna("").map(clean_text).tolist()
    headers = df.iloc[header_row].fillna("").map(clean_text).tolist()

    # 向右填滿科目（合併儲存格）
    fixed = []
    last = ""
    for s in subjects:
        if s != "":
            last = s
        fixed.append(last)
    subjects = fixed

    seat_idx = None
    name_idx = None
    for j, h in enumerate(headers):
        if seat_idx is None and "座號" in h:
            seat_idx = j
        if name_idx is None and "姓名" in h:
            name_idx = j

    if seat_idx is None:
        raise ValueError("找不到『座號』欄（帳號=座號 模式需要）。")
    if name_idx is None:
        raise ValueError("找不到『姓名』欄。")

    data = df.iloc[header_row + 1:].copy().fillna("")
    data = data.applymap(clean_text)

    # 移除空列
    data = data[data.apply(lambda r: any(str(x).strip() != "" for x in r), axis=1)]
    # 姓名空的列也移除
    data = data[data.iloc[:, name_idx].astype(str).str.strip() != ""]

    return df, data, subjects, evals, seat_idx, name_idx


# ===================== 學生視圖資料 =====================
@dataclass
class StudentView:
    seat: str
    name: str
    scores_df: pd.DataFrame  # 科目, 評量範圍, 分數, 分數數字


def build_student_view(data, subjects, evals, seat_idx, name_idx, seat_value: str) -> StudentView:
    target = None
    for _, row in data.iterrows():
        if seat_to_str(row.iloc[seat_idx]) == seat_value:
            target = row
            break

    if target is None:
        raise ValueError(f"查不到座號 {seat_value} 的資料。")

    name = clean_text(target.iloc[name_idx])

    rows = []
    n_cols = data.shape[1]
    for j in range(n_cols):
        if j in (seat_idx, name_idx):
            continue

        sval = clean_text(target.iloc[j])
        if is_hidden_score(sval):
            continue

        subj = subjects[j] if j < len(subjects) else ""
        rng = evals[j] if j < len(evals) else ""
        num = to_float_or_none(sval)

        rows.append({
            "科目": subj if subj else "-",
            "評量範圍": rng if rng else f"第{j+1}欄",
            "分數": sval,
            "分數數字": num
        })

    if not rows:
        raise ValueError("你這一列沒有任何成績欄位資料。")

    return StudentView(seat=seat_value, name=name, scores_df=pd.DataFrame(rows))


def build_student_view_by_row(data, subjects, evals, seat_idx, name_idx, row) -> StudentView:
    seat_value = seat_to_str(row.iloc[seat_idx])
    name = clean_text(row.iloc[name_idx])

    rows = []
    n_cols = data.shape[1]
    for j in range(n_cols):
        if j in (seat_idx, name_idx):
            continue

        sval = clean_text(row.iloc[j])
        if is_hidden_score(sval):
            continue

        subj = subjects[j] if j < len(subjects) else ""
        rng = evals[j] if j < len(evals) else ""
        num = to_float_or_none(sval)

        rows.append({
            "科目": subj if subj else "-",
            "評量範圍": rng if rng else f"第{j+1}欄",
            "分數": sval,
            "分數數字": num
        })

    if not rows:
        rows = [{"科目": "-", "評量範圍": "-", "分數": "-", "分數數字": None}]

    return StudentView(seat=seat_value, name=name, scores_df=pd.DataFrame(rows))


# ===================== 班級平均 =====================
def compute_class_avg(data, subjects, evals, seat_idx, name_idx):
    n_cols = data.shape[1]
    bucket = {}

    for _, row in data.iterrows():
        for j in range(n_cols):
            if j in (seat_idx, name_idx):
                continue

            sval = clean_text(row.iloc[j])
            if is_hidden_score(sval):
                continue

            num = to_float_or_none(sval)
            if num is None:
                continue

            subj = subjects[j] if j < len(subjects) else ""
            if subj == "":
                subj = "-"
            bucket.setdefault(subj, []).append(num)

    out = []
    for subj, arr in bucket.items():
        out.append({"科目": subj, "班級平均": sum(arr) / len(arr), "樣本數": len(arr)})

    if out:
        return pd.DataFrame(out).sort_values("科目")
    return pd.DataFrame(columns=["科目", "班級平均", "樣本數"])


# ===================== 排名 =====================
def compute_student_overall_avg(row, seat_idx, name_idx):
    nums = []
    for j in range(len(row)):
        if j in (seat_idx, name_idx):
            continue

        sval = clean_text(row.iloc[j])
        if is_hidden_score(sval):
            continue

        num = to_float_or_none(sval)
        if num is None:
            continue

        nums.append(num)

    if not nums:
        return None, 0
    return sum(nums) / len(nums), len(nums)


def compute_class_ranking(data, seat_idx, name_idx):
    rows = []
    for _, r in data.iterrows():
        seat = seat_to_str(r.iloc[seat_idx])
        name = clean_text(r.iloc[name_idx])
        if seat == "" or name == "":
            continue

        avg, n = compute_student_overall_avg(r, seat_idx, name_idx)
        rows.append({
            "座號": seat,
            "姓名": name,
            "平均": avg,
            "可計算筆數": n
        })

    ranking = pd.DataFrame(rows)
    if ranking.empty:
        return ranking.assign(名次=pd.Series(dtype=int), 百分位=pd.Series(dtype=float))

    has_avg = ranking["平均"].notna()
    ranked = ranking[has_avg].copy()
    ranked["名次"] = ranked["平均"].rank(ascending=False, method="min").astype(int)

    n_people = len(ranked)
    if n_people == 1:
        ranked["百分位"] = 100.0
    else:
        ranked["百分位"] = (1 - (ranked["名次"] - 1) / (n_people - 1)) * 100

    out = ranking.merge(ranked[["座號", "名次", "百分位"]], on="座號", how="left")
    out = out.sort_values("名次", na_position="last").reset_index(drop=True)
    return out


# ===================== PDF 輔助 =====================
def score_label(subj: str, eval_name: str) -> str:
    subj = clean_text(subj)
    eval_name = clean_text(eval_name)
    parts = [x for x in [subj, eval_name] if x]
    return "｜".join(parts) if parts else "-"


def compute_label_benchmarks(data, subjects, evals, seat_idx, name_idx):
    out = {}
    n_cols = data.shape[1]
    for j in range(n_cols):
        if j in (seat_idx, name_idx):
            continue
        vals = []
        for _, row in data.iterrows():
            num = to_float_or_none(row.iloc[j])
            if num is not None:
                vals.append(num)
        if not vals:
            continue
        s = pd.Series(vals, dtype=float)
        label = score_label(subjects[j] if j < len(subjects) else "", evals[j] if j < len(evals) else "")
        out[label] = {
            "頂": round(float(s.quantile(0.88)), 1),
            "前": round(float(s.quantile(0.75)), 1),
            "均": round(float(s.quantile(0.50)), 1),
            "後": round(float(s.quantile(0.25)), 1),
            "底": round(float(s.quantile(0.12)), 1),
        }
    return out


def build_benchmark_table(student: StudentView, data, subjects, evals, seat_idx, name_idx):
    if data is None:
        return Spacer(1, 0.1 * cm)

    bm_map = compute_label_benchmarks(data, subjects, evals, seat_idx, name_idx)
    rows = [["項目", "頂", "前", "均", "後", "底"]]
    for _, r in student.scores_df.iterrows():
        label = score_label(r.get("科目", ""), r.get("評量範圍", ""))
        if label not in bm_map:
            continue
        b = bm_map[label]
        short_label = clean_text(r.get("評量範圍", "")) or clean_text(r.get("科目", "")) or label
        rows.append([short_label, b["頂"], b["前"], b["均"], b["後"], b["底"]])

    if len(rows) == 1:
        return Spacer(1, 0.1 * cm)

    max_rows = 8
    rows = rows[: max_rows + 1]
    tbl = Table(rows, colWidths=[4.8 * cm, 0.75 * cm, 0.75 * cm, 0.75 * cm, 0.75 * cm, 0.75 * cm])
    style = [
        ("FONTNAME", (0, 0), (-1, -1), FONT),
        ("FONTSIZE", (0, 0), (-1, -1), 7.2),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAEAEA")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.grey),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]
    for i in range(1, len(rows)):
        if i % 2 == 1:
            style.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F8F8F8")))
    tbl.setStyle(TableStyle(style))
    return tbl


def build_compare_df(student: StudentView, data, subjects, evals, seat_idx, name_idx):
    if data is None:
        return pd.DataFrame(columns=["科目", "班級平均", "個人平均"])

    class_avg = compute_class_avg(data, subjects, evals, seat_idx, name_idx)
    mine_num = student.scores_df.dropna(subset=["分數數字"]).copy()
    if mine_num.empty:
        return pd.DataFrame(columns=["科目", "班級平均", "個人平均"])

    mine_by_subj = (
        mine_num.groupby("科目", as_index=False)["分數數字"]
        .mean()
        .rename(columns={"分數數字": "個人平均"})
    )
    class_avg2 = (
        class_avg.groupby("科目", as_index=False)["班級平均"].mean()
        if not class_avg.empty else pd.DataFrame(columns=["科目", "班級平均"])
    )
    compare = pd.merge(class_avg2, mine_by_subj, on="科目", how="inner")
    compare["班級平均"] = pd.to_numeric(compare["班級平均"], errors="coerce")
    compare["個人平均"] = pd.to_numeric(compare["個人平均"], errors="coerce")
    compare = compare.dropna(subset=["班級平均", "個人平均"], how="all")
    return compare


def make_compare_chart_image(student: StudentView, data, subjects, evals, seat_idx, name_idx):
    compare = build_compare_df(student, data, subjects, evals, seat_idx, name_idx)
    if compare.empty:
        return Spacer(1, 0.1 * cm)

    compare = compare.sort_values("科目").reset_index(drop=True)
    labels = compare["科目"].astype(str).tolist()[:6]
    class_vals = [float(v) if pd.notna(v) else 0.0 for v in compare["班級平均"].tolist()[:6]]
    mine_vals = [float(v) if pd.notna(v) else 0.0 for v in compare["個人平均"].tolist()[:6]]

    width = 8.1 * cm
    height = 5.0 * cm
    d = Drawing(width, height)

    title = String(4, height - 9, "全班平均 vs 個人平均", fontName=FONT, fontSize=7.5)
    d.add(title)

    chart = VerticalBarChart()
    chart.x = 20
    chart.y = 16
    chart.width = width - 32
    chart.height = height - 34
    chart.data = [class_vals, mine_vals]
    chart.strokeColor = colors.grey
    chart.valueAxis.valueMin = 0
    chart.valueAxis.valueMax = max(100, int(max(class_vals + mine_vals + [100]) / 10 + 1) * 10)
    chart.valueAxis.valueStep = 20
    chart.valueAxis.labels.fontName = FONT
    chart.valueAxis.labels.fontSize = 6
    chart.categoryAxis.categoryNames = labels
    chart.categoryAxis.labels.fontName = FONT
    chart.categoryAxis.labels.fontSize = 6
    chart.categoryAxis.labels.boxAnchor = 'n'
    chart.categoryAxis.labels.dy = -2
    chart.barWidth = 7
    chart.groupSpacing = 8
    chart.barSpacing = 2
    chart.bars[0].fillColor = colors.HexColor('#BFBFBF')
    chart.bars[1].fillColor = colors.HexColor('#7F7F7F')
    d.add(chart)

    d.add(String(width - 68, height - 9, '班級平均', fontName=FONT, fontSize=6.5, fillColor=colors.HexColor('#555555')))
    d.add(String(width - 30, height - 9, '個人平均', fontName=FONT, fontSize=6.5, fillColor=colors.black))
    return d


def build_student_report_story(student: StudentView, title_text: str, data=None, subjects=None, evals=None, seat_idx=None, name_idx=None):
    base_styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "BigTitle", parent=base_styles["Title"],
        fontName=FONT, fontSize=22, leading=26,
        alignment=1, spaceAfter=6
    )
    info_style = ParagraphStyle(
        "Info", parent=base_styles["Normal"],
        fontName=FONT, fontSize=12, leading=15
    )

    scores_df = student.scores_df.copy()
    numeric = scores_df["分數數字"].dropna().tolist()
    avg = (sum(numeric) / len(numeric)) if numeric else None

    row_count = len(scores_df)
    table_font = 13 if row_count <= 14 else 12 if row_count <= 18 else 11
    table_lead_pad = 5 if row_count <= 18 else 4

    story = []
    story.append(Paragraph(f"{title_text} 成績單", title_style))
    story.append(Spacer(1, 0.12 * cm))

    extra_avg = f"　平均：{avg:.1f} 分" if avg is not None else ""
    info_text = f"姓名：{student.name}  座號：{student.seat}{extra_avg}"
    info_table = Table([[Paragraph(info_text, info_style)]], colWidths=[18.2 * cm])
    info_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.8, colors.grey),
        ("BACKGROUND", (0, 0), (-1, -1), colors.whitesmoke),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 0.28 * cm))

    table_rows = [["科目", "評量範圍", "分數"]] + scores_df[["科目", "評量範圍", "分數"]].values.tolist()
    main_table = Table(table_rows, colWidths=[4.2 * cm, 10.8 * cm, 3.2 * cm])
    style_cmds = [
        ("FONTNAME", (0, 0), (-1, -1), FONT),
        ("FONTSIZE", (0, 0), (-1, -1), table_font),
        ("BOX", (0, 0), (-1, -1), 1, colors.black),
        ("GRID", (0, 0), (-1, -1), 0.45, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("ALIGN", (1, 1), (1, -1), "LEFT"),
        ("ALIGN", (2, 1), (2, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), table_lead_pad),
        ("TOPPADDING", (0, 0), (-1, -1), table_lead_pad),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]
    for r in range(1, len(table_rows)):
        if r % 2 == 1:
            style_cmds.append(("BACKGROUND", (0, r), (-1, r), colors.HexColor("#FAFAFA")))
    main_table.setStyle(TableStyle(style_cmds))
    story.append(main_table)
    story.append(Spacer(1, 0.16 * cm))

    bench_table = build_benchmark_table(student, data, subjects, evals, seat_idx, name_idx)
    chart_img = make_compare_chart_image(student, data, subjects, evals, seat_idx, name_idx)
    bottom = Table([[bench_table, chart_img]], colWidths=[8.8 * cm, 8.8 * cm])
    bottom.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(bottom)
    return story


# ===================== PDF：單一學生 =====================
def make_single_student_pdf_bytes(student: StudentView, title_text: str, data=None, subjects=None, evals=None, seat_idx=None, name_idx=None):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        rightMargin=1.1 * cm, leftMargin=1.1 * cm,
        topMargin=1.0 * cm, bottomMargin=1.0 * cm
    )
    story = build_student_report_story(student, title_text, data, subjects, evals, seat_idx, name_idx)
    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# ===================== PDF：全班（students list） =====================
def make_class_pdf_from_students(students: list, title_text: str, data=None, subjects=None, evals=None, seat_idx=None, name_idx=None):
    story = []
    for i, student in enumerate(students):
        story.extend(build_student_report_story(student, title_text, data, subjects, evals, seat_idx, name_idx))
        if i != len(students) - 1:
            story.append(PageBreak())

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        rightMargin=1.1 * cm, leftMargin=1.1 * cm,
        topMargin=1.0 * cm, bottomMargin=1.0 * cm
    )
    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# ===================== Streamlit UI =====================
st.set_page_config(page_title="成績查詢系統", layout="centered")
st.title("📌 成績查詢系統")

# ---- 讀 config.yaml ----
if not os.path.exists(CONFIG_PATH):
    st.error("找不到 config.yaml（請確認與 321.py 在同一資料夾）")
    st.stop()

with open(CONFIG_PATH, "r", encoding="utf-8") as f:
    config = yaml.safe_load(f)

authenticator = stauth.Authenticate(
    config["credentials"],
    config["cookie"]["name"],
    config["cookie"]["key"],
    config["cookie"]["expiry_days"],
    auto_hash=True
)

authenticator.login(location="main")

auth_status = st.session_state.get("authentication_status", None)
display_name = st.session_state.get("name", None)
username = st.session_state.get("username", None)

if auth_status is False:
    st.error("帳號或密碼錯誤")
    st.stop()
if auth_status is None:
    st.info("請先登入")
    st.stop()

authenticator.logout("登出", "sidebar")

role = config["credentials"]["usernames"].get(username, {}).get("role", "student")
st.sidebar.success(f"已登入：{display_name}（帳號：{username}｜身分：{role}）")

store = load_store()
if store:
    st.caption(f"📦 資料版本：{store['meta'].get('version','-')}｜更新時間：{store['meta'].get('updated_at','-')}")
else:
    st.caption("📦 目前尚未上傳成績資料")

st.divider()


# ===================== Admin / Student 分流 =====================
if role == "admin":
    st.subheader("🛠️ 老師/管理者：更新成績資料")

    uploaded = st.file_uploader("上傳成績 Excel（.xlsx/.xls）", type=["xlsx", "xls"])

    c1, c2, c3 = st.columns(3)
    with c1:
        subject_row = st.number_input("科目列（0-based）", min_value=0, value=DEFAULT_SUBJECT_ROW, step=1)
    with c2:
        eval_row = st.number_input("評量範圍列（0-based）", min_value=0, value=DEFAULT_EVAL_ROW, step=1)
    with c3:
        header_row = st.number_input("欄名列（含座號/姓名）（0-based）", min_value=0, value=DEFAULT_HEADER_ROW, step=1)

    title_text = st.text_input("成績標題（例如：小考/期中/模考）", value="小考")

    if uploaded:
        excel_bytes = uploaded.read()
        xls = pd.ExcelFile(io.BytesIO(excel_bytes))
        sheet_name = st.selectbox("選工作表", xls.sheet_names)

        if st.button("✅ 解析並保存（讓全班可查）"):
            try:
                _, data_admin, subjects_admin, evals_admin, seat_idx_admin, name_idx_admin = parse_all_scores_from_bytes(
                    excel_bytes, sheet_name, int(subject_row), int(eval_row), int(header_row)
                )

                meta = {
                    "version": sha256_hex(excel_bytes),
                    "updated_at": now_taipei_str(),
                    "sheet": sheet_name,
                    "subject_row": int(subject_row),
                    "eval_row": int(eval_row),
                    "header_row": int(header_row),
                    "title_text": title_text,
                    "rows": int(len(data_admin)),
                }
                save_store({"excel_bytes": excel_bytes, "meta": meta})

                append_log({
                    "time": meta["updated_at"],
                    "event": "admin_update_ok",
                    "username": username,
                    "msg": f"sheet={sheet_name}, rows={meta['rows']}, version={meta['version']}",
                })

                st.success("✅ 已更新！學生重新整理就能看到最新成績。")
                with st.expander("預覽前 5 列"):
                    st.dataframe(data_admin.head(5), use_container_width=True)

            except Exception as e:
                append_log({
                    "time": now_taipei_str(),
                    "event": "admin_update_failed",
                    "username": username,
                    "msg": str(e),
                })
                st.error(f"❌ 更新失敗：{e}")

    st.divider()
    st.subheader("📤 管理者匯出")

    store2 = load_store()
    if store2 is None:
        st.info("尚未有成績資料，請先上傳 Excel。")
        st.stop()

    excel_bytes2 = store2["excel_bytes"]
    meta2 = store2["meta"]

    excel_filename = f"original_{meta2.get('title_text','scores')}_{meta2.get('updated_at','')}.xlsx".replace(":", "-")
    st.download_button(
        "⬇️ 下載原始 Excel（管理者限定）",
        data=excel_bytes2,
        file_name=excel_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # 解析一次，下面多處共用
    try:
        _, data2, subjects2, evals2, seat_idx2, name_idx2 = parse_all_scores_from_bytes(
            excel_bytes2,
            meta2["sheet"],
            meta2["subject_row"],
            meta2["eval_row"],
            meta2["header_row"]
        )
    except Exception as e:
        st.error(f"❌ 系統資料解析失敗：{e}")
        st.stop()

    st.divider()
    st.subheader("🏆 全班排名（管理者限定）")
    try:
        ranking_df2 = compute_class_ranking(data2, seat_idx2, name_idx2)
        if ranking_df2.empty:
            st.info("目前沒有可排名資料（可能沒有任何數字分數）。")
        else:
            st.dataframe(
                ranking_df2[["名次", "座號", "姓名", "平均", "可計算筆數", "百分位"]],
                use_container_width=True
            )

            csv_bytes = ranking_df2.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(
                "⬇️ 下載排名 CSV",
                data=csv_bytes,
                file_name=f"ranking_{meta2.get('title_text','scores')}_{meta2.get('updated_at','')}.csv".replace(":", "-"),
                mime="text/csv"
            )
    except Exception as e:
        st.error(f"❌ 排名計算失敗：{e}")

    st.divider()
    st.subheader("📄 全班 PDF（管理者限定）")

    if st.button("📄 產生全班成績單 PDF（單一檔案）"):
        try:
            rows_list = []
            for _, r in data2.iterrows():
                seat = seat_to_str(r.iloc[seat_idx2])
                if seat != "":
                    rows_list.append(r)

            rows_list.sort(key=lambda r: seat_to_int_safe(seat_to_str(r.iloc[seat_idx2])))

            students = [
                build_student_view_by_row(data2, subjects2, evals2, seat_idx2, name_idx2, r)
                for r in rows_list
            ]

            class_pdf = make_class_pdf_from_students(students, title_text=meta2.get("title_text", "成績"), data=data2, subjects=subjects2, evals=evals2, seat_idx=seat_idx2, name_idx=name_idx2)
            pdf_name = f"class_scores_{meta2.get('title_text','scores')}_{meta2.get('updated_at','')}.pdf".replace(":", "-")

            st.download_button(
                "⬇️ 下載全班 PDF",
                data=class_pdf,
                file_name=pdf_name,
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"❌ 產生全班 PDF 失敗：{e}")

else:
    # ===================== 學生模式（只給 student 看） =====================
    st.subheader("📄 我的成績")

    store = load_store()
    if store is None:
        st.info("等待老師/管理者上傳成績。")
        st.stop()

    excel_bytes = store["excel_bytes"]
    meta = store["meta"]
    seat_value = clean_text(username)  # 本系統設定：帳號=座號

    try:
        _, data, subjects, evals, seat_idx, name_idx = parse_all_scores_from_bytes(
            excel_bytes, meta["sheet"], meta["subject_row"], meta["eval_row"], meta["header_row"]
        )
    except Exception as e:
        st.error(f"系統資料解析失敗：{e}")
        st.stop()

    all_seats = sorted(
        {seat_to_str(x) for x in data.iloc[:, seat_idx].tolist() if seat_to_str(x) != ""},
        key=lambda x: seat_to_int_safe(x)
    )

    if seat_value not in all_seats:
        st.error("❌ 查不到你的座號資料")
        st.info(
            "可能原因：\n"
            "- Excel 的座號欄有空格或格式不同（例如 01 vs 1）\n"
            "- 你登入的帳號不是座號（本系統設定：帳號=座號）\n\n"
            "建議：請老師確認 Excel『座號』欄格式，或把你的帳號改成座號。"
        )
        append_log({
            "time": now_taipei_str(),
            "event": "student_not_found",
            "username": username,
            "msg": f"seat_value={seat_value} not in sheet",
        })
        st.stop()

    try:
        student = build_student_view(data, subjects, evals, seat_idx, name_idx, seat_value)
    except Exception as e:
        st.error(f"❌ 顯示失敗：{e}")
        append_log({
            "time": now_taipei_str(),
            "event": "student_view_failed",
            "username": username,
            "msg": str(e),
        })
        st.stop()

    st.success(f"你好，{student.name}（座號 {student.seat}）")

    # ===== 排名（學生只看自己的名次，安全版）=====
    try:
        ranking_df = compute_class_ranking(data, seat_idx, name_idx)
        me = ranking_df[ranking_df["座號"] == student.seat]
        if len(me) == 1 and pd.notna(me.iloc[0]["名次"]):
            my_rank = int(me.iloc[0]["名次"])
            my_avg = float(me.iloc[0]["平均"])
            my_pct = float(me.iloc[0]["百分位"])
            total_ranked = int(ranking_df["名次"].dropna().max()) if ranking_df["名次"].notna().any() else 0

            st.info(f"🏅 你的總平均：{my_avg:.1f}｜名次：第 {my_rank} 名（共 {total_ranked} 人可排名）｜百分位：約 {my_pct:.0f}%")

            with st.expander("📌 名次附近（你前後各 2 名）", expanded=False):
                nearby = ranking_df[ranking_df["名次"].between(my_rank - 2, my_rank + 2, inclusive="both")].copy()
                st.dataframe(nearby[["名次", "座號", "姓名", "平均", "可計算筆數"]], use_container_width=True)
        else:
            st.info("🏅 目前沒有足夠的『數字分數』可計算總平均與排名（可能都是缺考/免試/文字）。")
    except Exception as e:
        st.warning(f"排名計算暫時無法顯示：{e}")

    st.dataframe(student.scores_df[["科目", "評量範圍", "分數"]], use_container_width=True)

    with st.expander("📊 分析與圖表（你 vs 班平均）", expanded=True):
        class_avg = compute_class_avg(data, subjects, evals, seat_idx, name_idx)

        mine_num = student.scores_df.dropna(subset=["分數數字"]).copy()
        if len(mine_num) == 0:
            st.info("你目前沒有可計算的數字分數（可能都是缺考/免試/文字）。")
        else:
            mine_by_subj = (
                mine_num.groupby("科目", as_index=False)["分數數字"]
                .mean()
                .rename(columns={"分數數字": "我的平均"})
            )

            class_avg2 = (
                class_avg.groupby("科目", as_index=False)["班級平均"].mean()
                if not class_avg.empty
                else pd.DataFrame(columns=["科目", "班級平均"])
            )

            compare = pd.merge(class_avg2, mine_by_subj, on="科目", how="outer")
            compare["班級平均"] = pd.to_numeric(compare["班級平均"], errors="coerce")
            compare["我的平均"] = pd.to_numeric(compare["我的平均"], errors="coerce")
            compare = compare.dropna(subset=["班級平均", "我的平均"], how="all")

            st.dataframe(compare, use_container_width=True)

            if compare.empty:
                st.warning("目前沒有可用的數字資料可以畫圖（班平均/我的平均可能都是空或非數字）。")
            else:
                line_df = compare.melt(
                    id_vars=["科目"],
                    value_vars=["班級平均", "我的平均"],
                    var_name="類別",
                    value_name="分數"
                ).dropna()

                chart = (
                    alt.Chart(line_df)
                    .mark_line(point=True)
                    .encode(
                        x=alt.X("科目:N", title=None),
                        y=alt.Y("分數:Q", title="分數"),
                        color=alt.Color("類別:N", legend=alt.Legend(title=None)),
                        tooltip=["科目:N", "類別:N", alt.Tooltip("分數:Q", format=".1f")]
                    )
                    .properties(height=320)
                )
                st.altair_chart(chart, use_container_width=True)

    pdf_bytes = make_single_student_pdf_bytes(student, title_text=meta.get("title_text", "成績"), data=data, subjects=subjects, evals=evals, seat_idx=seat_idx, name_idx=name_idx)
    st.download_button(
        "⬇️ 下載我的 PDF 成績單",
        data=pdf_bytes,
        file_name=f"score_{seat_value}.pdf",
        mime="application/pdf"
    )
