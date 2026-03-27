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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


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


def normalize_store(store):
    """向下相容舊版 store，統一轉成多考試格式"""
    if not store:
        return {"schema_version": 2, "active_exam": None, "exams": {}}

    if isinstance(store, dict) and "exams" in store:
        store.setdefault("schema_version", 2)
        store.setdefault("active_exam", next(iter(store["exams"]), None))
        return store

    if isinstance(store, dict) and "excel_bytes" in store and "meta" in store:
        exam_id = store["meta"].get("version", "legacy_exam")
        return {
            "schema_version": 2,
            "active_exam": exam_id,
            "exams": {
                exam_id: {
                    "excel_bytes": store["excel_bytes"],
                    "meta": store["meta"],
                    "analysis_config": {
                        "avg_fields": [],
                        "benchmark_fields": [],
                        "compare_fields": [],
                    },
                }
            },
        }

    return {"schema_version": 2, "active_exam": None, "exams": {}}


def load_store():
    if not os.path.exists(STORE_PATH):
        return normalize_store(None)
    with open(STORE_PATH, "rb") as f:
        raw = pickle.load(f)
    return normalize_store(raw)


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

    data = data[data.apply(lambda r: any(str(x).strip() != "" for x in r), axis=1)]
    data = data[data.iloc[:, name_idx].astype(str).str.strip() != ""]

    return df, data, subjects, evals, seat_idx, name_idx


def get_score_columns(subjects, evals, seat_idx, name_idx, n_cols):
    cols = []
    for j in range(n_cols):
        if j in (seat_idx, name_idx):
            continue
        subj = subjects[j] if j < len(subjects) else ""
        rng = evals[j] if j < len(evals) else ""
        label_parts = [x for x in [subj, rng] if clean_text(x) != ""]
        label = "｜".join(label_parts) if label_parts else f"第{j+1}欄"
        cols.append({
            "index": j,
            "subject": subj if subj else "-",
            "eval": rng,
            "label": label
        })
    return cols


def infer_default_analysis_fields(data, score_columns):
    selected = []
    for col in score_columns:
        idx = col["index"]
        values = data.iloc[:, idx].astype(str).map(clean_text).tolist()
        numeric_count = sum(to_float_or_none(v) is not None for v in values if not is_hidden_score(v))
        if numeric_count == 0:
            continue

        label = col["label"]
        low_keywords = ["單字", "作業", "平時", "閱讀", "默寫", "聽寫", "抽背", "小考", "習作", "訂正"]
        if any(k in label for k in low_keywords):
            continue
        selected.append(label)

    if not selected:
        selected = [c["label"] for c in score_columns]
    return selected


def labels_to_indices(score_columns, labels):
    label_set = set(labels or [])
    return [c["index"] for c in score_columns if c["label"] in label_set]


def make_analysis_config(score_columns, avg_labels, benchmark_labels, compare_labels, pdf_include_benchmarks=False):
    return {
        "avg_fields": avg_labels or [],
        "benchmark_fields": benchmark_labels or [],
        "compare_fields": compare_labels or [],
        "avg_indices": labels_to_indices(score_columns, avg_labels),
        "benchmark_indices": labels_to_indices(score_columns, benchmark_labels),
        "compare_indices": labels_to_indices(score_columns, compare_labels),
        "pdf_include_benchmarks": bool(pdf_include_benchmarks),
    }


def choose_indices(preferred, fallback, seat_idx, name_idx, row_len):
    if preferred:
        return preferred
    return [j for j in range(row_len) if j not in (seat_idx, name_idx) and j in fallback]


# ===================== 學生視圖資料 =====================
@dataclass
class StudentView:
    seat: str
    name: str
    scores_df: pd.DataFrame


def build_student_view(data, subjects, evals, seat_idx, name_idx, seat_value: str):
    target = None
    for _, row in data.iterrows():
        if seat_to_str(row.iloc[seat_idx]) == seat_value:
            target = row
            break

    if target is None:
        raise ValueError(f"查不到座號 {seat_value} 的資料。")

    return build_student_view_by_row(data, subjects, evals, seat_idx, name_idx, target)


def build_student_view_by_row(data, subjects, evals, seat_idx, name_idx, row):
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
            "欄位索引": j,
            "科目": subj if subj else "-",
            "評量範圍": rng if rng else f"第{j+1}欄",
            "欄位標籤": "｜".join([x for x in [subj, rng] if clean_text(x) != ""]) or f"第{j+1}欄",
            "分數": sval,
            "分數數字": num
        })

    if not rows:
        raise ValueError("你這一列沒有任何成績欄位資料。")

    return StudentView(seat=seat_value, name=name, scores_df=pd.DataFrame(rows))


# ===================== 統計 =====================
def compute_class_avg(data, subjects, evals, seat_idx, name_idx, allowed_indices=None):
    n_cols = data.shape[1]
    bucket = {}

    for _, row in data.iterrows():
        for j in range(n_cols):
            if j in (seat_idx, name_idx):
                continue
            if allowed_indices is not None and j not in allowed_indices:
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


def compute_student_overall_avg(row, seat_idx, name_idx, allowed_indices=None):
    nums = []
    for j in range(len(row)):
        if j in (seat_idx, name_idx):
            continue
        if allowed_indices is not None and j not in allowed_indices:
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


def compute_class_ranking(data, seat_idx, name_idx, allowed_indices=None):
    rows = []
    for _, r in data.iterrows():
        seat = seat_to_str(r.iloc[seat_idx])
        name = clean_text(r.iloc[name_idx])
        if seat == "" or name == "":
            continue

        avg, n = compute_student_overall_avg(r, seat_idx, name_idx, allowed_indices=allowed_indices)
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


def calc_benchmarks(scores: pd.Series):
    scores = pd.to_numeric(scores, errors="coerce").dropna()
    if scores.empty:
        return None
    return {
        "頂標": round(scores.quantile(0.88), 1),
        "前標": round(scores.quantile(0.75), 1),
        "均標": round(scores.quantile(0.50), 1),
        "後標": round(scores.quantile(0.25), 1),
        "底標": round(scores.quantile(0.12), 1),
        "樣本數": int(len(scores)),
    }


def compute_benchmark_table(data, subjects, evals, seat_idx, name_idx, benchmark_indices=None):
    if not benchmark_indices:
        return pd.DataFrame(columns=["欄位", "頂標", "前標", "均標", "後標", "底標", "樣本數"])

    rows = []
    n_cols = data.shape[1]
    for j in range(n_cols):
        if j in (seat_idx, name_idx) or j not in benchmark_indices:
            continue
        label_parts = []
        subj = subjects[j] if j < len(subjects) else ""
        rng = evals[j] if j < len(evals) else ""
        if clean_text(subj):
            label_parts.append(subj)
        if clean_text(rng):
            label_parts.append(rng)
        label = "｜".join(label_parts) if label_parts else f"第{j+1}欄"

        scores = data.iloc[:, j].map(clean_text).map(to_float_or_none)
        bench = calc_benchmarks(scores)
        if bench:
            row = {"欄位": label}
            row.update(bench)
            rows.append(row)

    if not rows:
        return pd.DataFrame(columns=["欄位", "頂標", "前標", "均標", "後標", "底標", "樣本數"])
    return pd.DataFrame(rows)


def filter_benchmarks_for_student(bench_df, student_scores_df):
    if bench_df.empty or student_scores_df.empty:
        return bench_df.iloc[0:0].copy()

    own_labels = set(student_scores_df["欄位標籤"].dropna().astype(str).tolist())
    own_subjects = set(student_scores_df["科目"].dropna().astype(str).tolist())

    mask = bench_df["欄位"].astype(str).isin(own_labels) | bench_df["欄位"].astype(str).isin(own_subjects)
    return bench_df[mask].copy()


def get_exam_choices(store):
    exams = store.get("exams", {})
    items = []
    for exam_id, exam in exams.items():
        meta = exam.get("meta", {})
        label = f"{meta.get('exam_name', meta.get('title_text', exam_id))}｜{meta.get('updated_at', '-')}"
        items.append((exam_id, label))
    items.sort(key=lambda x: x[1], reverse=True)
    return items


def load_exam_dataset(exam_obj):
    excel_bytes = exam_obj["excel_bytes"]
    meta = exam_obj["meta"]
    _, data, subjects, evals, seat_idx, name_idx = parse_all_scores_from_bytes(
        excel_bytes,
        meta["sheet"],
        meta["subject_row"],
        meta["eval_row"],
        meta["header_row"]
    )
    score_columns = get_score_columns(subjects, evals, seat_idx, name_idx, data.shape[1])

    analysis_config = exam_obj.get("analysis_config", {})
    if not analysis_config.get("avg_indices") and analysis_config.get("avg_fields") is not None:
        analysis_config["avg_indices"] = labels_to_indices(score_columns, analysis_config.get("avg_fields", []))
    if not analysis_config.get("benchmark_indices") and analysis_config.get("benchmark_fields") is not None:
        analysis_config["benchmark_indices"] = labels_to_indices(score_columns, analysis_config.get("benchmark_fields", []))
    if not analysis_config.get("compare_indices") and analysis_config.get("compare_fields") is not None:
        analysis_config["compare_indices"] = labels_to_indices(score_columns, analysis_config.get("compare_fields", []))

    return {
        "excel_bytes": excel_bytes,
        "meta": meta,
        "data": data,
        "subjects": subjects,
        "evals": evals,
        "seat_idx": seat_idx,
        "name_idx": name_idx,
        "score_columns": score_columns,
        "analysis_config": analysis_config,
    }


def build_compare_table(student_view: StudentView, compare_indices):
    if student_view.scores_df.empty:
        return pd.DataFrame(columns=["科目", "分數"])
    df = student_view.scores_df.copy()
    if compare_indices:
        df = df[df["欄位索引"].isin(compare_indices)].copy()
    if df.empty:
        return pd.DataFrame(columns=["科目", "分數"])
    return df[["科目", "分數數字"]].dropna().groupby("科目", as_index=False)["分數數字"].mean().rename(columns={"分數數字": "分數"})


# ===================== PDF：單一學生 =====================
def make_single_student_pdf_bytes(student: StudentView, title_text: str, student_bench_df=None, include_benchmarks=False):
    base_styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "BigTitle", parent=base_styles["Title"],
        fontName=FONT, fontSize=20, leading=24,
        alignment=1, spaceAfter=8
    )
    info_style = ParagraphStyle(
        "Info", parent=base_styles["Normal"],
        fontName=FONT, fontSize=11, leading=14, spaceAfter=4
    )
    summary_style = ParagraphStyle(
        "Summary", parent=base_styles["Normal"],
        fontName=FONT, fontSize=10, leading=13
    )

    scores_df = student.scores_df.copy()
    numeric = scores_df["分數數字"].dropna().tolist()
    avg = (sum(numeric) / len(numeric)) if numeric else None
    mx = max(numeric) if numeric else None
    mn = min(numeric) if numeric else None

    story = []
    story.append(Paragraph(f"{title_text} 成績單", title_style))
    story.append(Spacer(1, 0.2 * cm))

    extra_avg = f"　平均：{avg:.1f} 分" if avg is not None else ""
    info_text = f"姓名：{student.name}　座號：{student.seat}{extra_avg}"
    info_table = Table([[Paragraph(info_text, info_style)]], colWidths=[18 * cm])
    info_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.8, colors.grey),
        ("BACKGROUND", (0, 0), (-1, -1), colors.whitesmoke),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 0.4 * cm))

    table_rows = [["科目", "評量範圍", "分數"]] + scores_df[["科目", "評量範圍", "分數"]].values.tolist()
    table = Table(table_rows, colWidths=[4.0 * cm, 10.0 * cm, 2.5 * cm])
    style_cmds = [
        ("FONTNAME", (0, 0), (-1, -1), FONT),
        ("FONTSIZE", (0, 0), (-1, -1), 11),
        ("BOX", (0, 0), (-1, -1), 1, colors.black),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("ALIGN", (1, 1), (1, -1), "LEFT"),
        ("ALIGN", (2, 1), (2, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
    ]
    for r in range(1, len(table_rows)):
        if r % 2 == 1:
            style_cmds.append(("BACKGROUND", (0, r), (-1, r), colors.HexColor("#F7F7F7")))
    table.setStyle(TableStyle(style_cmds))
    story.append(table)

    story.append(Spacer(1, 0.3 * cm))
    if numeric:
        lines = [
            f"‧ 共有 {len(numeric)} 筆可計算成績（只計算數字分數）",
            f"‧ 最高分：{mx:.1f}",
            f"‧ 最低分：{mn:.1f}",
            f"‧ 平均分：{avg:.1f}",
        ]
    else:
        lines = ["‧ 沒有可計算的數字分數（可能都是缺考/免試/文字）"]

    story.append(Paragraph("<br/>".join(lines), summary_style))

    if include_benchmarks and student_bench_df is not None and not student_bench_df.empty:
        story.append(Spacer(1, 0.3 * cm))
        story.append(Paragraph("頂前均後底標", info_style))
        bench_rows = [["欄位", "頂標", "前標", "均標", "後標", "底標"]]
        for _, r in student_bench_df.iterrows():
            bench_rows.append([
                str(r.get("欄位", "")),
                str(r.get("頂標", "")),
                str(r.get("前標", "")),
                str(r.get("均標", "")),
                str(r.get("後標", "")),
                str(r.get("底標", "")),
            ])
        bench_table = Table(bench_rows, colWidths=[6.0 * cm, 2.2 * cm, 2.2 * cm, 2.2 * cm, 2.2 * cm, 2.2 * cm])
        bench_table.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), FONT),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("BOX", (0, 0), (-1, -1), 0.8, colors.black),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        story.append(bench_table)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        rightMargin=1.5 * cm, leftMargin=1.5 * cm,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm
    )
    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# ===================== PDF：全班（students list） =====================
def make_class_pdf_from_students(students: list, title_text: str, benchmark_map=None, include_benchmarks=False):
    base_styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "BigTitle", parent=base_styles["Title"],
        fontName=FONT, fontSize=20, leading=24,
        alignment=1, spaceAfter=8
    )
    info_style = ParagraphStyle(
        "Info", parent=base_styles["Normal"],
        fontName=FONT, fontSize=11, leading=14, spaceAfter=4
    )
    summary_style = ParagraphStyle(
        "Summary", parent=base_styles["Normal"],
        fontName=FONT, fontSize=10, leading=13
    )

    story = []

    for i, student in enumerate(students):
        scores_df = student.scores_df.copy()
        numeric = scores_df["分數數字"].dropna().tolist()
        avg = (sum(numeric) / len(numeric)) if numeric else None
        mx = max(numeric) if numeric else None
        mn = min(numeric) if numeric else None

        story.append(Paragraph(f"{title_text} 成績單", title_style))
        story.append(Spacer(1, 0.2 * cm))

        extra_avg = f"　平均：{avg:.1f} 分" if avg is not None else ""
        info_text = f"姓名：{student.name}　座號：{student.seat}{extra_avg}"
        info_table = Table([[Paragraph(info_text, info_style)]], colWidths=[18 * cm])
        info_table.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.8, colors.grey),
            ("BACKGROUND", (0, 0), (-1, -1), colors.whitesmoke),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(info_table)
        story.append(Spacer(1, 0.4 * cm))

        table_rows = [["科目", "評量範圍", "分數"]] + scores_df[["科目", "評量範圍", "分數"]].values.tolist()
        table = Table(table_rows, colWidths=[4.0 * cm, 10.0 * cm, 2.5 * cm])
        style_cmds = [
            ("FONTNAME", (0, 0), (-1, -1), FONT),
            ("FONTSIZE", (0, 0), (-1, -1), 11),
            ("BOX", (0, 0), (-1, -1), 1, colors.black),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ("ALIGN", (1, 1), (1, -1), "LEFT"),
            ("ALIGN", (2, 1), (2, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
        ]
        for r in range(1, len(table_rows)):
            if r % 2 == 1:
                style_cmds.append(("BACKGROUND", (0, r), (-1, r), colors.HexColor("#F7F7F7")))
        table.setStyle(TableStyle(style_cmds))
        story.append(table)

        story.append(Spacer(1, 0.3 * cm))
        if numeric:
            lines = [
                f"‧ 共有 {len(numeric)} 筆可計算成績（只計算數字分數）",
                f"‧ 最高分：{mx:.1f}",
                f"‧ 最低分：{mn:.1f}",
                f"‧ 平均分：{avg:.1f}",
            ]
        else:
            lines = ["‧ 沒有可計算的數字分數（可能都是缺考/免試/文字）"]
        story.append(Paragraph("<br/>".join(lines), summary_style))

        student_bench_df = None
        if benchmark_map:
            student_bench_df = benchmark_map.get(student.seat)

        if include_benchmarks and student_bench_df is not None and not student_bench_df.empty:
            story.append(Spacer(1, 0.3 * cm))
            story.append(Paragraph("頂前均後底標", info_style))
            bench_rows = [["欄位", "頂標", "前標", "均標", "後標", "底標"]]
            for _, r in student_bench_df.iterrows():
                bench_rows.append([
                    str(r.get("欄位", "")),
                    str(r.get("頂標", "")),
                    str(r.get("前標", "")),
                    str(r.get("均標", "")),
                    str(r.get("後標", "")),
                    str(r.get("底標", "")),
                ])
            bench_table = Table(bench_rows, colWidths=[6.0 * cm, 2.2 * cm, 2.2 * cm, 2.2 * cm, 2.2 * cm, 2.2 * cm])
            bench_table.setStyle(TableStyle([
                ("FONTNAME", (0, 0), (-1, -1), FONT),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("BOX", (0, 0), (-1, -1), 0.8, colors.black),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]))
            story.append(bench_table)

        if i != len(students) - 1:
            story.append(PageBreak())

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        rightMargin=1.5 * cm, leftMargin=1.5 * cm,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm
    )
    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# ===================== Streamlit UI =====================
st.set_page_config(page_title="成績查詢系統", layout="centered")
st.title("📌 成績查詢系統")

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
exam_choices = get_exam_choices(store)

if exam_choices:
    active_exam = store.get("active_exam") or exam_choices[0][0]
    active_meta = store["exams"][active_exam]["meta"]
    st.caption(f"📦 目前預設考試：{active_meta.get('exam_name', active_meta.get('title_text', active_exam))}｜更新時間：{active_meta.get('updated_at','-')}")
else:
    st.caption("📦 目前尚未上傳成績資料")

st.divider()


if role == "admin":
    st.subheader("🛠️ 老師/管理者：新增或更新段考資料")

    uploaded = st.file_uploader("上傳成績 Excel（.xlsx/.xls）", type=["xlsx", "xls"])

    c1, c2, c3 = st.columns(3)
    with c1:
        subject_row = st.number_input("科目列（0-based）", min_value=0, value=DEFAULT_SUBJECT_ROW, step=1)
    with c2:
        eval_row = st.number_input("評量範圍列（0-based）", min_value=0, value=DEFAULT_EVAL_ROW, step=1)
    with c3:
        header_row = st.number_input("欄名列（含座號/姓名）（0-based）", min_value=0, value=DEFAULT_HEADER_ROW, step=1)

    exam_name = st.text_input("考試名稱（例如：高二下第一次段考）", value="高二下第一次段考")
    title_text = st.text_input("PDF / 顯示標題（例如：第一次段考）", value="第一次段考")

    if uploaded:
        excel_bytes = uploaded.read()
        xls = pd.ExcelFile(io.BytesIO(excel_bytes))
        sheet_name = st.selectbox("選工作表", xls.sheet_names)

        try:
            _, data_admin, subjects_admin, evals_admin, seat_idx_admin, name_idx_admin = parse_all_scores_from_bytes(
                excel_bytes, sheet_name, int(subject_row), int(eval_row), int(header_row)
            )
            score_columns_admin = get_score_columns(
                subjects_admin, evals_admin, seat_idx_admin, name_idx_admin, data_admin.shape[1]
            )
            all_labels = [c["label"] for c in score_columns_admin]
            default_main = infer_default_analysis_fields(data_admin, score_columns_admin)

            st.markdown("### ✅ 分析欄位設定")
            avg_labels = st.multiselect(
                "哪些欄位要列入平均 / 排名",
                all_labels,
                default=default_main
            )
            benchmark_labels = st.multiselect(
                "哪些欄位要列入五標",
                all_labels,
                default=default_main
            )
            compare_labels = st.multiselect(
                "哪些欄位要列入歷次比較",
                all_labels,
                default=default_main
            )
            pdf_include_benchmarks = st.checkbox("成績單 PDF 要列印頂前均後底標", value=False)

            with st.expander("預覽前 5 列", expanded=False):
                st.dataframe(data_admin.head(5), use_container_width=True)

            if st.button("✅ 保存這份考試資料"):
                exam_id = sha256_hex(excel_bytes + exam_name.encode("utf-8"))
                meta = {
                    "version": sha256_hex(excel_bytes),
                    "updated_at": now_taipei_str(),
                    "sheet": sheet_name,
                    "subject_row": int(subject_row),
                    "eval_row": int(eval_row),
                    "header_row": int(header_row),
                    "title_text": title_text,
                    "exam_name": exam_name,
                    "rows": int(len(data_admin)),
                }
                analysis_config = make_analysis_config(
                    score_columns_admin, avg_labels, benchmark_labels, compare_labels,
                    pdf_include_benchmarks=pdf_include_benchmarks
                )

                store = load_store()
                store["exams"][exam_id] = {
                    "excel_bytes": excel_bytes,
                    "meta": meta,
                    "analysis_config": analysis_config,
                }
                store["active_exam"] = exam_id
                save_store(store)

                append_log({
                    "time": meta["updated_at"],
                    "event": "admin_save_exam_ok",
                    "username": username,
                    "msg": f"exam={exam_name}, sheet={sheet_name}, rows={meta['rows']}, version={meta['version']}",
                })

                st.success("✅ 已保存！現在這份考試也會出現在學生端可選清單。")

        except Exception as e:
            append_log({
                "time": now_taipei_str(),
                "event": "admin_update_failed",
                "username": username,
                "msg": str(e),
            })
            st.error(f"❌ 解析失敗：{e}")

    st.divider()
    st.subheader("📚 已保存的考試資料")

    store = load_store()
    exam_choices = get_exam_choices(store)

    if not exam_choices:
        st.info("尚未有任何考試資料。")
        st.stop()

    exam_id_to_label = dict(exam_choices)
    active_exam_id = store.get("active_exam") or exam_choices[0][0]
    selected_admin_exam = st.selectbox(
        "選擇要管理 / 匯出的考試",
        options=[eid for eid, _ in exam_choices],
        format_func=lambda x: exam_id_to_label.get(x, x),
        index=[eid for eid, _ in exam_choices].index(active_exam_id) if active_exam_id in [eid for eid, _ in exam_choices] else 0
    )

    if st.button("設成學生端預設考試"):
        store["active_exam"] = selected_admin_exam
        save_store(store)
        st.success("✅ 已更新預設考試。")

    exam_obj = store["exams"][selected_admin_exam]
    dataset = load_exam_dataset(exam_obj)
    meta2 = dataset["meta"]
    data2 = dataset["data"]
    subjects2 = dataset["subjects"]
    evals2 = dataset["evals"]
    seat_idx2 = dataset["seat_idx"]
    name_idx2 = dataset["name_idx"]
    analysis_config2 = dataset["analysis_config"]

    st.caption(
        f"考試名稱：{meta2.get('exam_name','-')}｜資料筆數：{meta2.get('rows','-')}｜更新時間：{meta2.get('updated_at','-')}｜PDF列印五標：{'是' if analysis_config2.get('pdf_include_benchmarks') else '否'}"
    )

    excel_filename = f"original_{meta2.get('exam_name','scores')}_{meta2.get('updated_at','')}.xlsx".replace(":", "-")
    st.download_button(
        "⬇️ 下載這份考試原始 Excel",
        data=dataset["excel_bytes"],
        file_name=excel_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.divider()
    st.subheader("🏆 全班排名（依管理者勾選欄位）")
    try:
        ranking_df2 = compute_class_ranking(
            data2,
            seat_idx2,
            name_idx2,
            allowed_indices=analysis_config2.get("avg_indices", [])
        )
        if ranking_df2.empty:
            st.info("目前沒有可排名資料。")
        else:
            st.dataframe(
                ranking_df2[["名次", "座號", "姓名", "平均", "可計算筆數", "百分位"]],
                use_container_width=True
            )
    except Exception as e:
        st.error(f"❌ 排名計算失敗：{e}")

    st.divider()
    st.subheader("📏 五標（依管理者勾選欄位）")
    bench_df = compute_benchmark_table(
        data2, subjects2, evals2, seat_idx2, name_idx2,
        benchmark_indices=analysis_config2.get("benchmark_indices", [])
    )
    if bench_df.empty:
        st.info("這份考試目前沒有五標欄位。")
    else:
        st.dataframe(bench_df, use_container_width=True)

    st.divider()
    st.subheader("📄 全班 PDF（管理者限定）")
    if st.button("📄 產生這份考試的全班成績單 PDF"):
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
            bench_full_df = compute_benchmark_table(
                data2, subjects2, evals2, seat_idx2, name_idx2,
                benchmark_indices=analysis_config2.get("benchmark_indices", [])
            )
            benchmark_map = {}
            for stu in students:
                benchmark_map[stu.seat] = filter_benchmarks_for_student(bench_full_df, stu.scores_df)

            class_pdf = make_class_pdf_from_students(
                students,
                title_text=meta2.get("title_text", "成績"),
                benchmark_map=benchmark_map,
                include_benchmarks=analysis_config2.get("pdf_include_benchmarks", False)
            )
            pdf_name = f"class_scores_{meta2.get('exam_name','scores')}_{meta2.get('updated_at','')}.pdf".replace(":", "-")

            st.download_button(
                "⬇️ 下載全班 PDF",
                data=class_pdf,
                file_name=pdf_name,
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"❌ 產生全班 PDF 失敗：{e}")

else:
    st.subheader("📄 我的成績")

    store = load_store()
    exam_choices = get_exam_choices(store)

    if not exam_choices:
        st.info("等待老師/管理者上傳成績。")
        st.stop()

    exam_id_to_label = dict(exam_choices)
    active_exam_id = store.get("active_exam") or exam_choices[0][0]
    selected_exam = st.selectbox(
        "選擇考試",
        options=[eid for eid, _ in exam_choices],
        format_func=lambda x: exam_id_to_label.get(x, x),
        index=[eid for eid, _ in exam_choices].index(active_exam_id) if active_exam_id in [eid for eid, _ in exam_choices] else 0
    )

    selected_exam_obj = store["exams"][selected_exam]
    ds = load_exam_dataset(selected_exam_obj)

    seat_value = clean_text(username)

    data = ds["data"]
    subjects = ds["subjects"]
    evals = ds["evals"]
    seat_idx = ds["seat_idx"]
    name_idx = ds["name_idx"]
    meta = ds["meta"]
    analysis_config = ds["analysis_config"]

    all_seats = sorted(
        {seat_to_str(x) for x in data.iloc[:, seat_idx].tolist() if seat_to_str(x) != ""},
        key=lambda x: seat_to_int_safe(x)
    )

    if seat_value not in all_seats:
        st.error("❌ 查不到你的座號資料")
        st.info(
            "可能原因：\n"
            "- Excel 的座號欄有空格或格式不同（例如 01 vs 1）\n"
            "- 你登入的帳號不是座號（本系統設定：帳號=座號）"
        )
        st.stop()

    try:
        student = build_student_view(data, subjects, evals, seat_idx, name_idx, seat_value)
    except Exception as e:
        st.error(f"❌ 顯示失敗：{e}")
        st.stop()

    st.success(f"你好，{student.name}（座號 {student.seat}）")
    st.caption(f"目前查看：{meta.get('exam_name', meta.get('title_text', '-'))}")

    try:
        ranking_df = compute_class_ranking(
            data, seat_idx, name_idx,
            allowed_indices=analysis_config.get("avg_indices", [])
        )
        me = ranking_df[ranking_df["座號"] == student.seat]
        if len(me) == 1 and pd.notna(me.iloc[0]["名次"]):
            my_rank = int(me.iloc[0]["名次"])
            my_avg = float(me.iloc[0]["平均"])
            my_pct = float(me.iloc[0]["百分位"])
            total_ranked = int(ranking_df["名次"].dropna().max()) if ranking_df["名次"].notna().any() else 0

            st.info(f"🏅 你的分析平均：{my_avg:.1f}｜名次：第 {my_rank} 名（共 {total_ranked} 人可排名）｜百分位：約 {my_pct:.0f}%")
        else:
            st.info("🏅 目前沒有足夠的分析欄位可計算平均與排名。")
    except Exception as e:
        st.warning(f"排名計算暫時無法顯示：{e}")

    st.dataframe(student.scores_df[["科目", "評量範圍", "分數"]], use_container_width=True)

    with st.expander("📊 分析與圖表（你 vs 班平均）", expanded=True):
        class_avg = compute_class_avg(
            data, subjects, evals, seat_idx, name_idx,
            allowed_indices=analysis_config.get("avg_indices", [])
        )

        mine_num = student.scores_df.dropna(subset=["分數數字"]).copy()
        if analysis_config.get("avg_indices"):
            mine_num = mine_num[mine_num["欄位索引"].isin(analysis_config.get("avg_indices", []))]

        if len(mine_num) == 0:
            st.info("你目前沒有被納入分析的數字分數。")
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

            if not compare.empty:
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

    with st.expander("📈 與其他段考比較", expanded=False):
        compare_candidates = [eid for eid, _ in exam_choices if eid != selected_exam]
        if not compare_candidates:
            st.info("目前只有一份考試資料，還不能比較。")
        else:
            compare_exam = st.selectbox(
                "選擇要比較的另一份考試",
                options=[""] + compare_candidates,
                format_func=lambda x: "請選擇" if x == "" else exam_id_to_label.get(x, x)
            )
            if compare_exam:
                ds_old = load_exam_dataset(store["exams"][compare_exam])
                data_old = ds_old["data"]
                seat_idx_old = ds_old["seat_idx"]
                name_idx_old = ds_old["name_idx"]

                all_old_seats = {seat_to_str(x) for x in data_old.iloc[:, seat_idx_old].tolist() if seat_to_str(x) != ""}
                if seat_value not in all_old_seats:
                    st.warning("另一份考試裡查不到你的座號，無法比較。")
                else:
                    student_old = build_student_view(data_old, ds_old["subjects"], ds_old["evals"], seat_idx_old, name_idx_old, seat_value)

                    old_table = build_compare_table(student_old, ds_old["analysis_config"].get("compare_indices", []))
                    new_table = build_compare_table(student, analysis_config.get("compare_indices", []))

                    merged = pd.merge(
                        old_table, new_table, on="科目", how="outer", suffixes=(
                            f"（{ds_old['meta'].get('exam_name', '舊')})",
                            f"（{meta.get('exam_name', '新')})"
                        )
                    )
                    col_old = [c for c in merged.columns if c.startswith("分數（")][0]
                    col_new = [c for c in merged.columns if c.endswith(f"（{meta.get('exam_name', '新')})")][0]
                    merged["變化"] = pd.to_numeric(merged[col_new], errors="coerce") - pd.to_numeric(merged[col_old], errors="coerce")
                    st.dataframe(merged, use_container_width=True)

    with st.expander("📏 這份考試的五標", expanded=False):
        bench_df = compute_benchmark_table(
            data, subjects, evals, seat_idx, name_idx,
            benchmark_indices=analysis_config.get("benchmark_indices", [])
        )
        student_bench_df = filter_benchmarks_for_student(bench_df, student.scores_df)
        if student_bench_df.empty:
            st.info("老師這份考試沒有設定屬於你這次成績的五標欄位。")
        else:
            st.dataframe(student_bench_df, use_container_width=True)

    bench_df_for_pdf = compute_benchmark_table(
        data, subjects, evals, seat_idx, name_idx,
        benchmark_indices=analysis_config.get("benchmark_indices", [])
    )
    student_bench_df_for_pdf = filter_benchmarks_for_student(bench_df_for_pdf, student.scores_df)

    pdf_bytes = make_single_student_pdf_bytes(
        student,
        title_text=meta.get("title_text", "成績"),
        student_bench_df=student_bench_df_for_pdf,
        include_benchmarks=analysis_config.get("pdf_include_benchmarks", False)
    )
    st.download_button(
        "⬇️ 下載我的 PDF 成績單",
        data=pdf_bytes,
        file_name=f"score_{seat_value}_{meta.get('exam_name','exam')}.pdf".replace(":", "-"),
        mime="application/pdf"
    )
