"""Clean the FTIR extract, compute KPIs, and write one PowerPoint per caller."""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
DATA: pd.DataFrame | None = None
SOURCE: Path | None = None

RGB = {
    "navy": RGBColor(0x0B, 0x3A, 0x5B),
    "teal": RGBColor(0x1F, 0x7A, 0x8C),
    "amber": RGBColor(0xE0, 0x8A, 0x2A),
    "green": RGBColor(0x1E, 0x8C, 0x5A),
    "red": RGBColor(0xC0, 0x39, 0x2B),
    "slate": RGBColor(0x5D, 0x6D, 0x7E),
    "white": RGBColor(0xFF, 0xFF, 0xFF),
    "ink": RGBColor(0x1C, 0x28, 0x33),
    "paper": RGBColor(0xF4, 0xF7, 0xF8),
    "mist": RGBColor(0xD6, 0xE0, 0xE6),
}


def find_dataset() -> Path:
    names = ("ftir_dataset.csv", "ftir_dataset.db", "records.csv")
    for base in (ROOT, *ROOT.parents):
        for name in names:
            candidate = base / "data" / name
            if candidate.is_file():
                return candidate
    raise FileNotFoundError("No FTIR extract in data/. Expected ftir_dataset.csv, ftir_dataset.db, or records.csv.")


def load() -> pd.DataFrame:
    global DATA, SOURCE
    path = find_dataset()
    if path.suffix.lower() == ".db":
        import sqlite3

        with sqlite3.connect(path) as con:
            tables = pd.read_sql(
                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'",
                con,
            )
            if tables.empty:
                raise ValueError(f"No tables in {path}")
            chosen = tables["name"].iloc[0]
            for name in tables["name"].tolist():
                cols = pd.read_sql(f"SELECT * FROM {name} LIMIT 0", con).columns
                if "ftir_report_date" in cols:
                    chosen = name
                    break
            frame = pd.read_sql(f"SELECT * FROM {chosen}", con)
    else:
        frame = pd.read_csv(path, low_memory=False)
    DATA = clean(frame)
    SOURCE = path
    return DATA


def parse_km(value):
    if pd.isna(value):
        return pd.NA
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return int(round(value))
    text = str(value).lower().replace(",", "").strip().split("/")[0]
    text = " ".join(text.replace("km.", " ").replace("km", " ").split()).rstrip(".")
    if text in {"", "nan", "none"}:
        return pd.NA
    if text.endswith("k"):
        return int(round(float(text[:-1]) * 1000))
    return int(round(float(text)))


def km_bucket(km) -> str:
    if pd.isna(km):
        return "Unknown"
    km = int(km)
    if km <= 5000:
        return "0–5000"
    upper = ((km - 1) // 5000 + 1) * 5000
    lower = upper - 4999
    return f"{lower}–{upper}"


def bucket_lower(label: str) -> int:
    if label == "40001+":
        return 10**9
    if label == "Unknown":
        return -1
    return int(str(label).split("–")[0])


def chart_km_bucket(label: str) -> str:
    if label == "Unknown":
        return "Unknown"
    if bucket_lower(label) >= 40001:
        return "40001+"
    return label


def month_chart_label(months) -> str:
    if pd.isna(months):
        return "Unknown"
    months = int(months)
    return "25+" if months >= 25 else str(months)


def clean(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw.copy()
    df["using_km_int"] = df["using_time_km"].map(parse_km).astype("Int64")
    df["usage_bucket"] = df["using_km_int"].map(km_bucket)
    df["usage_bucket_chart"] = df["usage_bucket"].map(chart_km_bucket)
    df["days_used"] = pd.to_numeric(df["days_used"], errors="coerce")
    df["months_used"] = (df["days_used"] // 30).astype("Int64")
    df["ftir_report_date"] = pd.to_datetime(df["ftir_report_date"], errors="coerce")
    df["judgement_date"] = pd.to_datetime(df["judgement_date"], errors="coerce")
    df["report_year"] = df["ftir_report_date"].dt.year.astype("Int64")
    df["report_month"] = df["ftir_report_date"].dt.month.astype("Int64")
    df["fc_ok"] = df["fc_ok"].astype("string").str.strip().str.upper()
    df["rank"] = df["rank"].astype("string").str.strip().str.upper()
    df["problem_solved"] = df["problem_solved"].astype("string").str.strip()
    df["is_resolved"] = df["problem_solved"].fillna("").str.lower().eq("yes").astype(int)
    df["problem_solved"] = df["problem_solved"].fillna("Unknown")
    df["investigation_lag_days"] = (df["judgement_date"] - df["ftir_report_date"]).dt.days
    df["plant"] = df["manufacturer_factory"].fillna("Unknown").astype(str).str.strip()
    df["c_measure"] = df["c_measure"].fillna("Not recorded").astype(str).str.strip()
    df["sales_dealer"] = df["sales_dealer"].fillna("Unknown").astype(str).str.strip()
    df["service_dealer"] = df["service_dealer"].fillna("Unknown").astype(str).str.strip()
    return df


def _insights(k: dict) -> list[str]:
    notes = []
    share, plant = k["top_failure_share"], k["top_failure_plant"]
    if share >= 0.30:
        notes.append(f"{plant} accounts for {share:.0%} of failures — systemic concentration, not a one-off dealer event.")
    else:
        notes.append(f"{plant} is the largest failure source at {share:.0%} of NG cases.")
    if k["share_gt_20k"] > k["share_le_10k"]:
        notes.append(
            f"High mileage buckets (>20,000 km) hold more of the file ({k['share_gt_20k']:.0%}) "
            f"than the 0–10,000 km bands ({k['share_le_10k']:.0%})."
        )
    else:
        notes.append(
            f"Complaints concentrate early in life: {k['share_le_10k']:.0%} sit in the 0–10,000 km buckets, "
            f"while mileage above 20,000 km accounts for {k['share_gt_20k']:.0%}."
        )
    early_rate, late_rate = k["resolution_0_12"], k["resolution_13_plus"]
    if late_rate < early_rate - 0.05:
        notes.append(f"Resolution rate drops after 12 months of usage ({early_rate:.0%} through month 12 vs {late_rate:.0%} afterward).")
    elif late_rate > early_rate + 0.05:
        notes.append(f"Resolution rate is higher after 12 months ({late_rate:.0%}) than in the first year ({early_rate:.0%}).")
    else:
        notes.append(
            f"Resolution rate holds near {early_rate:.0%} through 12 months and {late_rate:.0%} after that — "
            "months in service are not what separates closed from open cases."
        )
    notes.append(
        f"A countermeasure other than \"No action - within spec\" is recorded in {k['action_rate']:.0%} of cases, "
        f"and {k['action_solved']:.0%} of those are marked solved "
        f"(solved rate with no action: {k['no_action_solved']:.0%}; overall {k['solved_rate']:.0%})."
    )
    notes.append(f"Average investigation lag from report to judgement is {k['avg_lag']:.0f} days (median {k['median_lag']:.0f}, n={k['n_lag']:,}).")
    notes.append(
        f"Severity rank {k['top_rank']} is {k['top_rank_share']:.0%} of the file "
        f"(A {k['rank_shares'].get('A', 0):.0%}, B {k['rank_shares'].get('B', 0):.0%}, C {k['rank_shares'].get('C', 0):.0%})."
    )
    return notes


def _recommendations(k: dict) -> list[str]:
    rank_line = (
        f"Rank A is already {k['top_rank_share']:.0%} of the file. Review it on its own queue, not mixed with B and C."
        if k["top_rank"] == "A"
        else (
            f"Rank {k['top_rank']} is {k['top_rank_share']:.0%} of volume. Keep rank A "
            f"({k['rank_shares'].get('A', 0):.0%}) on a separate weekly list so it is not averaged into the majority."
        )
    )
    return [
        f"Start the plant review at {k['top_failure_plant']} ({k['top_failure_share']:.0%} of NG). Compare its top subjects and causal parts with the next plant before treating the gap as volume only.",
        "Put the early-usage population first: most FTIR volume is under 10,000 km, so Pareto the subjects inside the 0–5000 and 5001–10000 buckets.",
        "Solved rate barely moves between countermeasure types. Audit whether problem_solved is updated after the measure, and whether the measure matches the checked result.",
        f"Investigation lag averages {k['avg_lag']:.0f} days. A 30-day judgement target would cut about {max(k['avg_lag'] - 30, 0):.0f} days off the current mean.",
        rank_line,
    ]


def analyze(df: pd.DataFrame) -> dict:
    failures = df[df["fc_ok"] == "NG"]
    n = len(df)
    n_fail = len(failures)
    plant_counts = df["plant"].value_counts()
    plant_fail = failures["plant"].value_counts()
    plant_kpi = pd.DataFrame({"complaints": plant_counts, "failures_ng": plant_fail}).fillna(0).astype(int)
    plant_kpi["pct_of_failures"] = plant_kpi["failures_ng"] / n_fail if n_fail else 0
    plant_kpi = plant_kpi.sort_values("failures_ng", ascending=False)

    usage = (
        df.groupby("usage_bucket_chart").size().rename("complaints").reset_index()
    )
    usage["ord"] = usage["usage_bucket_chart"].map(bucket_lower)
    usage = usage.sort_values("ord")

    month_chart = (
        df.assign(month_label=df["months_used"].map(month_chart_label))
        .groupby("month_label").size().rename("complaints").reset_index()
    )
    month_chart["ord"] = month_chart["month_label"].map(lambda label: 25 if label == "25+" else int(label))
    month_chart = month_chart.sort_values("ord")

    rank_counts = df["rank"].value_counts()
    rank_rows = []
    for rank in ("A", "B", "C"):
        count = int(rank_counts.get(rank, 0))
        rank_rows.append({"rank": rank, "complaints": count, "pct": count / n if n else 0})
    top_rank = max(rank_rows, key=lambda row: row["complaints"])

    sales_counts = df["sales_dealer"].value_counts()
    service_counts = df["service_dealer"].value_counts()
    top_dealers = sales_counts.head(8).index.tolist()
    dealers = [
        {
            "dealer": dealer,
            "sales": int(sales_counts.get(dealer, 0)),
            "service": int(service_counts.get(dealer, 0)),
        }
        for dealer in top_dealers
    ]

    measure = pd.crosstab(df["c_measure"], df["problem_solved"], dropna=False)
    for col in ("Yes", "Partially", "No", "Unknown"):
        if col not in measure.columns:
            measure[col] = 0
    measure["complaints"] = measure[["Yes", "Partially", "No", "Unknown"]].sum(axis=1)
    measure = measure.sort_values("complaints", ascending=True)

    ok = df.loc[df["fc_ok"].eq("OK")].dropna(subset=["ftir_report_date"])
    monthly = (
        ok.assign(month_start=lambda frame: frame["ftir_report_date"].dt.to_period("M").dt.to_timestamp())
        .groupby("month_start").size().sort_index()
    )
    lag = df["investigation_lag_days"].dropna()
    lag = lag[lag >= 0]
    lag_labels = ["0–14", "15–29", "30–44", "45–59", "60–74", "75–90"]
    lag_cut = pd.cut(lag, bins=[0, 15, 30, 45, 60, 75, 91], right=False, labels=lag_labels)
    lag_chart = lag_cut.value_counts().reindex(lag_labels).fillna(0).astype(int)

    early = df.loc[df["months_used"] <= 12, "is_resolved"]
    late = df.loc[df["months_used"] > 12, "is_resolved"]
    action_mask = df["c_measure"] != "No action - within spec"
    top_plant = plant_kpi.index[0]
    date_min = df["ftir_report_date"].min()
    date_max = df["ftir_report_date"].max()
    kpi = {
        "n": n,
        "n_fail": n_fail,
        "fail_rate": n_fail / n if n else 0,
        "date_min": f"{date_min:%b %Y}",
        "date_max": f"{date_max:%b %Y}",
        "top_failure_plant": str(top_plant),
        "top_failure_share": float(plant_kpi.loc[top_plant, "pct_of_failures"]),
        "top_failure_count": int(plant_kpi.loc[top_plant, "failures_ng"]),
        "share_le_10k": float((df["using_km_int"] <= 10000).mean()),
        "share_gt_20k": float((df["using_km_int"] > 20000).mean()),
        "resolution_rate": float(df["is_resolved"].mean()),
        "resolution_0_12": float(early.mean()) if len(early) else 0,
        "resolution_13_plus": float(late.mean()) if len(late) else 0,
        "n_0_12": int(len(early)),
        "n_13_plus": int(len(late)),
        "top_rank": top_rank["rank"],
        "top_rank_share": top_rank["pct"],
        "rank_shares": {row["rank"]: row["pct"] for row in rank_rows},
        "top_sales_dealer": str(sales_counts.index[0]),
        "top_sales_count": int(sales_counts.iloc[0]),
        "top_service_dealer": str(service_counts.index[0]),
        "top_service_count": int(service_counts.iloc[0]),
        "same_dealer_share": float((df["sales_dealer"] == df["service_dealer"]).mean()),
        "action_rate": float(action_mask.mean()),
        "action_solved": float(df.loc[action_mask, "is_resolved"].mean()) if action_mask.any() else 0,
        "no_action_solved": float(df.loc[~action_mask, "is_resolved"].mean()) if (~action_mask).any() else 0,
        "solved_rate": float(df["is_resolved"].mean()),
        "avg_lag": float(lag.mean()) if len(lag) else 0,
        "median_lag": float(lag.median()) if len(lag) else 0,
        "n_lag": int(len(lag)),
        "ok_count": int(monthly.sum()),
    }
    charts = {
        "monthly": [{"label": label, "ok": int(value)} for label, value in zip(monthly.index.strftime("%Y-%b"), monthly.tolist())],
        "plants": [
            {
                "name": str(idx),
                "complaints": int(plant_kpi.loc[idx, "complaints"]),
                "failures": int(plant_kpi.loc[idx, "failures_ng"]),
                "fail_pct": float(plant_kpi.loc[idx, "pct_of_failures"]),
            }
            for idx in plant_kpi.sort_values("complaints", ascending=True).index
        ],
        "usage": [{"label": str(row.usage_bucket_chart), "complaints": int(row.complaints)} for row in usage.itertuples(index=False)],
        "months": [{"label": str(row.month_label), "complaints": int(row.complaints)} for row in month_chart.itertuples(index=False)],
        "ranks": rank_rows,
        "dealers": dealers,
        "measures": [
            {
                "name": str(idx),
                "yes": int(measure.loc[idx, "Yes"]),
                "partially": int(measure.loc[idx, "Partially"]),
                "no": int(measure.loc[idx, "No"]),
                "unknown": int(measure.loc[idx, "Unknown"]),
            }
            for idx in measure.index
        ],
        "lag": [{"label": str(label), "count": int(lag_chart.loc[label])} for label in lag_labels],
    }
    return {
        "kpis": kpi,
        "charts": charts,
        "insights": _insights(kpi),
        "recommendations": _recommendations(kpi),
        "_date_min": date_min,
        "_date_max": date_max,
        "_monthly_index_min": monthly.index.min(),
        "_monthly_index_max": monthly.index.max(),
    }


def _fill(shape, color: RGBColor) -> None:
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def _title_bar(slide, prs, title: str, subtitle: str | None = None) -> None:
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.92))
    _fill(bar, RGB["navy"])
    box = slide.shapes.add_textbox(Inches(0.45), Inches(0.14), Inches(12.4), Inches(0.38))
    paragraph = box.text_frame.paragraphs[0]
    paragraph.text = title
    paragraph.font.size = Pt(26)
    paragraph.font.bold = True
    paragraph.font.color.rgb = RGB["white"]
    paragraph.font.name = "Calibri"
    if subtitle:
        sub = slide.shapes.add_textbox(Inches(0.45), Inches(0.52), Inches(12.4), Inches(0.30))
        line = sub.text_frame.paragraphs[0]
        line.text = subtitle
        line.font.size = Pt(12)
        line.font.color.rgb = RGB["mist"]
        line.font.name = "Calibri"


def _bullets(slide, left, top, width, height, items, size=13) -> None:
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    frame = box.text_frame
    frame.word_wrap = True
    for index, item in enumerate(items):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = item
        paragraph.font.size = Pt(size)
        paragraph.font.name = "Calibri"
        paragraph.font.color.rgb = RGB["ink"]
        paragraph.space_after = Pt(6)


def _labels(chart, chart_type) -> None:
    plot = chart.plots[0]
    plot.has_data_labels = True
    labels = plot.data_labels
    labels.number_format = "#,##0;;"
    labels.font.name = "Calibri"
    stacked = chart_type in (XL_CHART_TYPE.BAR_STACKED, XL_CHART_TYPE.COLUMN_STACKED)
    if stacked:
        labels.position = XL_LABEL_POSITION.CENTER
        labels.font.size = Pt(8)
        labels.font.color.rgb = RGB["white"]
    elif chart_type in (XL_CHART_TYPE.LINE, XL_CHART_TYPE.LINE_MARKERS):
        labels.position = XL_LABEL_POSITION.ABOVE
        labels.font.size = Pt(7)
        labels.font.color.rgb = RGB["ink"]
    else:
        labels.position = XL_LABEL_POSITION.OUTSIDE_END
        labels.font.size = Pt(8)
        labels.font.color.rgb = RGB["ink"]


def _chart(slide, chart_type, categories, series, left, top, width, height, colors, legend=False, value_title=None, label_size=9, point_colors=None):
    data = CategoryChartData()
    data.categories = [str(item) for item in categories]
    for name, values in series:
        data.add_series(name, tuple(float(value) for value in values))
    chart = slide.shapes.add_chart(chart_type, Inches(left), Inches(top), Inches(width), Inches(height), data).chart
    chart.has_legend = legend
    if legend:
        chart.legend.include_in_layout = False
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(11)
        chart.legend.font.name = "Calibri"
    chart.has_title = False
    chart.category_axis.tick_labels.font.size = Pt(label_size)
    chart.category_axis.tick_labels.font.name = "Calibri"
    chart.category_axis.has_major_gridlines = False
    chart.value_axis.has_major_gridlines = True
    chart.value_axis.tick_labels.font.size = Pt(10)
    chart.value_axis.tick_labels.font.color.rgb = RGB["slate"]
    if value_title:
        chart.value_axis.has_title = True
        title = chart.value_axis.axis_title.text_frame.paragraphs[0]
        title.text = value_title
        title.font.size = Pt(10)
        title.font.color.rgb = RGB["slate"]
        title.font.name = "Calibri"
    for series_obj, color in zip(chart.series, colors):
        series_obj.format.fill.solid()
        series_obj.format.fill.fore_color.rgb = color
        series_obj.format.line.color.rgb = color
        if chart_type in (XL_CHART_TYPE.LINE, XL_CHART_TYPE.LINE_MARKERS):
            series_obj.format.line.width = Pt(2.25)
    if point_colors:
        for point, color in zip(chart.series[0].points, point_colors):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = color
    _labels(chart, chart_type)


def write_pptx(path: Path, name: str, result: dict) -> None:
    kpi = result["kpis"]
    charts = result["charts"]
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    slide = prs.slides.add_slide(blank)
    _fill(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height), RGB["navy"])
    _fill(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(4.55), prs.slide_width, Inches(0.08)), RGB["amber"])
    title = slide.shapes.add_textbox(Inches(0.7), Inches(1.9), Inches(11.5), Inches(1.3))
    line = title.text_frame.paragraphs[0]
    line.text = "FTIR Automated KPI Report – 2026"
    line.font.size = Pt(40)
    line.font.bold = True
    line.font.color.rgb = RGB["white"]
    line.font.name = "Calibri"
    sub = slide.shapes.add_textbox(Inches(0.7), Inches(3.4), Inches(11), Inches(0.9))
    frame = sub.text_frame
    frame.word_wrap = True
    prepared = frame.paragraphs[0]
    prepared.text = f"Prepared for {name}"
    prepared.font.size = Pt(20)
    prepared.font.color.rgb = RGB["white"]
    prepared.font.name = "Calibri"
    detail = frame.add_paragraph()
    detail.text = f"{kpi['n']:,} field reports  ·  {kpi['date_min']} – {kpi['date_max']}"
    detail.font.size = Pt(16)
    detail.font.color.rgb = RGB["mist"]
    detail.font.name = "Calibri"

    plants = charts["plants"]
    slide = prs.slides.add_slide(blank)
    _title_bar(slide, prs, "Monthly FC OK trend", f"Prepared for {name}  ·  fc_ok = OK only")
    _chart(
        slide, XL_CHART_TYPE.LINE_MARKERS,
        [row["label"] for row in charts["monthly"]],
        [("FC OK", [row["ok"] for row in charts["monthly"]])],
        0.35, 1.15, 9.15, 5.95, [RGB["green"]], value_title="FC OK count", label_size=8,
    )
    _bullets(slide, 9.65, 1.4, 3.35, 5.2, [
        f"FC OK rows: {kpi['ok_count']:,}.",
        f"Span {result['_monthly_index_min']:%Y-%b} to {result['_monthly_index_max']:%Y-%b}.",
        "NG rows are excluded. Each point is that report month.",
    ])

    slide = prs.slides.add_slide(blank)
    _title_bar(slide, prs, "Plant distribution", f"Prepared for {name}")
    _chart(
        slide, XL_CHART_TYPE.BAR_CLUSTERED,
        [row["name"] for row in plants],
        [("Complaints", [row["complaints"] for row in plants])],
        0.35, 1.15, 9.15, 5.95, [RGB["teal"]], value_title="FTIR count",
    )
    _bullets(slide, 9.65, 1.3, 3.35, 5.5, ["Share of failures (NG)"] + [
        f"{row['name']}: {row['fail_pct']:.0%} ({row['failures']:,})"
        for row in sorted(plants, key=lambda item: item["failures"], reverse=True)
    ], size=12)

    slide = prs.slides.add_slide(blank)
    _title_bar(slide, prs, "Usage buckets", f"Prepared for {name}  ·  5,000 km bands")
    _chart(
        slide, XL_CHART_TYPE.COLUMN_CLUSTERED,
        [row["label"] for row in charts["usage"]],
        [("Complaints", [row["complaints"] for row in charts["usage"]])],
        0.35, 1.15, 9.15, 5.95, [RGB["navy"]], value_title="FTIR count",
    )
    _bullets(slide, 9.65, 1.4, 3.35, 5.2, [
        f"0–10,000 km: {kpi['share_le_10k']:.0%} of complaints.",
        f"Above 20,000 km: {kpi['share_gt_20k']:.0%}.",
    ])

    slide = prs.slides.add_slide(blank)
    _title_bar(slide, prs, "Days used (months)", f"Prepared for {name}")
    _chart(
        slide, XL_CHART_TYPE.COLUMN_CLUSTERED,
        [row["label"] for row in charts["months"]],
        [("Complaints", [row["complaints"] for row in charts["months"]])],
        0.35, 1.15, 9.15, 5.95, [RGB["teal"]], value_title="FTIR count", label_size=9,
    )
    _bullets(slide, 9.65, 1.4, 3.35, 5.2, [
        f"Resolution, months 0–12: {kpi['resolution_0_12']:.0%} (n={kpi['n_0_12']:,}).",
        f"Resolution after month 12: {kpi['resolution_13_plus']:.0%} (n={kpi['n_13_plus']:,}).",
    ])

    rank_colors = {"A": RGB["red"], "B": RGB["amber"], "C": RGB["teal"]}
    slide = prs.slides.add_slide(blank)
    _title_bar(slide, prs, "Severity mix", f"Prepared for {name}")
    _chart(
        slide, XL_CHART_TYPE.COLUMN_CLUSTERED,
        [row["rank"] for row in charts["ranks"]],
        [("Complaints", [row["complaints"] for row in charts["ranks"]])],
        0.35, 1.15, 8.4, 5.95, [RGB["navy"]], value_title="FTIR count",
        point_colors=[rank_colors[row["rank"]] for row in charts["ranks"]],
    )
    _bullets(slide, 9.0, 1.6, 3.9, 4.5, [
        f"A: {kpi['rank_shares'].get('A', 0):.0%}",
        f"B: {kpi['rank_shares'].get('B', 0):.0%}",
        f"C: {kpi['rank_shares'].get('C', 0):.0%}",
    ], size=16)

    slide = prs.slides.add_slide(blank)
    _title_bar(slide, prs, "Dealer impact", f"Prepared for {name}")
    _chart(
        slide, XL_CHART_TYPE.COLUMN_CLUSTERED,
        [row["dealer"] for row in charts["dealers"]],
        [
            ("Sales dealer", [row["sales"] for row in charts["dealers"]]),
            ("Service dealer", [row["service"] for row in charts["dealers"]]),
        ],
        0.35, 1.15, 9.15, 5.95, [RGB["navy"], RGB["amber"]], legend=True, value_title="FTIR count", label_size=9,
    )
    sales_share = kpi["top_sales_count"] / kpi["n"] if kpi["n"] else 0
    spread = (
        "Counts are close across dealers; no single outlet dominates the file."
        if sales_share < 0.08
        else f"{kpi['top_sales_dealer']} alone is {sales_share:.0%} of sales-dealer complaints."
    )
    _bullets(slide, 9.65, 1.35, 3.35, 5.3, [
        f"Top sales: {kpi['top_sales_dealer']} ({kpi['top_sales_count']:,}).",
        f"Top service: {kpi['top_service_dealer']} ({kpi['top_service_count']:,}).",
        f"Same dealer on both roles: {kpi['same_dealer_share']:.0%}.",
        spread,
    ], size=13)

    slide = prs.slides.add_slide(blank)
    _title_bar(slide, prs, "Countermeasure effectiveness", f"Prepared for {name}")
    _chart(
        slide, XL_CHART_TYPE.BAR_STACKED,
        [row["name"] for row in charts["measures"]],
        [
            ("Yes", [row["yes"] for row in charts["measures"]]),
            ("Partially", [row["partially"] for row in charts["measures"]]),
            ("No", [row["no"] for row in charts["measures"]]),
            ("Unknown", [row["unknown"] for row in charts["measures"]]),
        ],
        0.3, 1.15, 9.2, 5.95,
        [RGB["green"], RGB["amber"], RGB["red"], RGB["slate"]],
        legend=True, value_title="FTIR count", label_size=10,
    )
    _bullets(slide, 9.65, 1.4, 3.35, 5.2, [
        f"Action recorded: {kpi['action_rate']:.0%} of rows.",
        f"Solved given an action: {kpi['action_solved']:.0%}.",
        f"Solved given no action: {kpi['no_action_solved']:.0%}.",
    ])

    slide = prs.slides.add_slide(blank)
    _title_bar(slide, prs, "Time to action", f"Prepared for {name}")
    _chart(
        slide, XL_CHART_TYPE.COLUMN_CLUSTERED,
        [row["label"] for row in charts["lag"]],
        [("FTIR count", [row["count"] for row in charts["lag"]])],
        0.35, 1.15, 9.15, 5.95, [RGB["navy"]], value_title="FTIR count",
    )
    _bullets(slide, 9.65, 1.4, 3.35, 5.2, [
        f"Mean lag {kpi['avg_lag']:.0f} days.",
        f"Median lag {kpi['median_lag']:.0f} days.",
        f"Rows with both dates: {kpi['n_lag']:,}.",
    ])

    slide = prs.slides.add_slide(blank)
    _title_bar(slide, prs, "Insights & Recommendations", f"Prepared for {name}")
    _fill(slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.35), Inches(1.15), Inches(6.25), Inches(5.95)), RGB["paper"])
    _fill(slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.75), Inches(1.15), Inches(6.2), Inches(5.95)), RGB["paper"])
    left_head = slide.shapes.add_textbox(Inches(0.55), Inches(1.28), Inches(5.9), Inches(0.34))
    left_head.text_frame.paragraphs[0].text = "Insights"
    left_head.text_frame.paragraphs[0].font.size = Pt(16)
    left_head.text_frame.paragraphs[0].font.bold = True
    left_head.text_frame.paragraphs[0].font.color.rgb = RGB["navy"]
    left_head.text_frame.paragraphs[0].font.name = "Calibri"
    right_head = slide.shapes.add_textbox(Inches(6.95), Inches(1.28), Inches(5.8), Inches(0.34))
    right_head.text_frame.paragraphs[0].text = "Recommendations"
    right_head.text_frame.paragraphs[0].font.size = Pt(16)
    right_head.text_frame.paragraphs[0].font.bold = True
    right_head.text_frame.paragraphs[0].font.color.rgb = RGB["navy"]
    right_head.text_frame.paragraphs[0].font.name = "Calibri"
    _bullets(slide, 0.55, 1.75, 5.85, 5.1, result["insights"], size=12)
    _bullets(slide, 6.95, 1.75, 5.8, 5.1, result["recommendations"], size=12)

    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(path)


REQUIRED_COLUMNS = (
    "using_time_km",
    "days_used",
    "ftir_report_date",
    "judgement_date",
    "fc_ok",
    "rank",
    "problem_solved",
    "manufacturer_factory",
    "c_measure",
    "sales_dealer",
    "service_dealer",
)


def frame_from_csv(path: Path) -> pd.DataFrame:
    frame = pd.read_csv(path, low_memory=False)
    missing = [column for column in REQUIRED_COLUMNS if column not in frame.columns]
    if missing:
        raise ValueError("Missing columns: " + ", ".join(missing))
    if frame.empty:
        raise ValueError("The uploaded file has no rows")
    return clean(frame)


def build_report(name: str, path: Path, frame: pd.DataFrame | None = None) -> dict:
    if frame is None:
        if DATA is None:
            load()
        frame = DATA
    result = analyze(frame)
    write_pptx(path, name, result)
    public = {
        "kpis": result["kpis"],
        "charts": result["charts"],
        "insights": result["insights"],
        "recommendations": result["recommendations"],
    }
    return public
