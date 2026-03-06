"""
SPSS Excel -> PowerPoint Report Generator
==========================================
Streamlit app: upload SPSS crosstab Excel, configure via Regie Tabel,
generate styled PowerPoint with python-pptx.
"""

import io
import re
from collections import OrderedDict

import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import (
    XL_CHART_TYPE,
    XL_LABEL_POSITION,
    XL_LEGEND_POSITION,
)
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN

# =====================================================================
# CONSTANTS
# =====================================================================
FONT_NAME = "Avenir Next LT Pro"
FONT_SIZE = Pt(10)

# Patterns in question+answer text that indicate a row should be SKIPPED
# entirely (these rows act as block separators — they trigger _flush).
SKIP_PATTERNS = [
    "comparisons of column proportions",
    "results are based on",
]

# Question-level skip: if the QUESTION text itself equals one of these,
# the entire question block is discarded.
SKIP_QUESTION_EXACT = {"topbox", "bottombox"}

# Answer-level utility rows: these are collected into the block (so we
# can extract n=) but are DROPPED from the chart data afterwards.
# IMPORTANT: do NOT put these in SKIP_ANSWER_EXACT — that would flush
# the block prematurely and orphan any rows that come after.
DROP_ANSWER_OPTIONS = {"n", "%", "topbox", "bottombox", "top2box", "bottom2box"}

STACKED_KEYWORDS = [
    "eens", "oneens", "goed", "slecht", "vaak", "nooit",
    "belangrijk", "tevreden", "ontevreden",
]
STACKED_PREFIXES = ["zeer ", "heel ", "helemaal "]

STACKED_COLOR_MAP = {
    "helemaal eens":          "#39B54A",
    "helemaal mee eens":      "#39B54A",
    "zeer goed":              "#39B54A",
    "zeer tevreden":          "#39B54A",
    "mee eens":               "#A8CD66",
    "eens":                   "#A8CD66",
    "goed":                   "#A8CD66",
    "tevreden":               "#A8CD66",
    "neutraal":               "#FFC04D",
    "niet eens, niet oneens": "#FFC04D",
    "niet eens/niet oneens":  "#FFC04D",
    "mee oneens":             "#F28E2B",
    "oneens":                 "#F28E2B",
    "slecht":                 "#F28E2B",
    "ontevreden":             "#F28E2B",
    "helemaal oneens":        "#E1001A",
    "helemaal mee oneens":    "#E1001A",
    "zeer slecht":            "#E1001A",
    "zeer ontevreden":        "#E1001A",
    "weet ik niet":           "#808080",
    "weet niet":              "#808080",
    "geen mening":            "#808080",
}

BAR_PRIMARY   = "#C60651"
BAR_SECONDARY = "#FF85A2"
BAR_GREY      = "#D3D3D3"

BOTTOM_LABELS = {
    "weet ik niet", "weet niet", "geen van bovenstaande", "geen mening",
}
PENULTIMATE_LABELS = {
    "anders, namelijk", "anders", "overig",
    "anders, namelijk...", "anders, namelijk:",
    "anders, namelijk …", "anders, namelijk…",
    "een ander netwerk, namelijk",
    "een ander netwerk, namelijk ...",
    "een ander netwerk, namelijk…",
}

SLIDE_WIDTH  = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

DARK_GREY = RGBColor(0x33, 0x33, 0x33)
MID_GREY  = RGBColor(0x80, 0x80, 0x80)


# =====================================================================
# HELPERS
# =====================================================================
def hex_to_rgb(hex_str: str) -> RGBColor:
    h = hex_str.lstrip("#")
    return RGBColor(int(h[:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _is_empty(val) -> bool:
    if val is None:
        return True
    if isinstance(val, float) and pd.isna(val):
        return True
    s = str(val).replace("\xa0", " ").strip().lower()
    return s in ("", "nan", "none")


def _clean_answer(raw: str) -> str:
    """Aggressively clean an answer string."""
    return (
        str(raw)
        .replace("\xa0", " ")
        .replace("\r\n", "\n")
        .replace("\r", "\n")
        .strip()
    )


def _is_utility_row(answer: str) -> bool:
    """Check if this is a utility row (n, %, topbox, bottombox, etc.).

    Handles multi-line values like 'n\\n%', hidden characters, padding.
    """
    clean = _clean_answer(answer).lower()

    # Exact match
    if clean in DROP_ANSWER_OPTIONS:
        return True

    # Multi-line: split on newlines and check if ALL parts are utility
    parts = [p.strip() for p in clean.split("\n") if p.strip()]
    if parts and all(p in DROP_ANSWER_OPTIONS for p in parts):
        return True

    # Combined like "n %" or "n%"
    if re.match(r'^[n%\s]+$', clean) and len(clean) <= 5:
        return True

    return False


def _row_has_n(answer: str) -> bool:
    """Check if this answer row represents an 'n' (sample size) row."""
    clean = _clean_answer(answer).lower()
    if clean == "n":
        return True
    # Multi-line: one of the parts is "n"
    parts = [p.strip() for p in clean.split("\n") if p.strip()]
    return "n" in parts


def _should_skip_row(q_cell: str, a_cell: str) -> bool:
    """Check if a row should be SKIPPED entirely (acts as block separator).

    Only use this for truly separating content like footnotes.
    Do NOT use this for n/% /topbox/bottombox — those must stay in the
    block so n= can be extracted before they are filtered out.
    """
    combined = (q_cell + " " + a_cell).lower().strip()
    for pat in SKIP_PATTERNS:
        if pat in combined:
            return True
    return False


def _set_font(font, size=None, bold=False, color=None, name=FONT_NAME):
    font.name = name
    font.size = size if size else FONT_SIZE
    if bold:
        font.bold = True
    if color:
        font.color.rgb = color


# =====================================================================
# 1. PARSE SPSS EXCEL
# =====================================================================
def parse_spss_excel(uploaded_file) -> tuple[OrderedDict, list[str]]:
    raw = pd.read_excel(uploaded_file, header=None, dtype=str)

    row0 = raw.iloc[0].fillna("").astype(str).str.strip()
    row1 = raw.iloc[1].fillna("").astype(str).str.strip()

    # Build column names from the two header rows
    col_names: list[str] = []
    current_main = ""
    # Characters that act as placeholder sub-headers (not real names)
    placeholder_subs = {"nan", "", "_", "-", "–", "—"}

    for idx, (main, sub) in enumerate(zip(row0, row1)):
        if main and main.lower() != "nan":
            current_main = main
        if sub and sub.lower() not in placeholder_subs:
            col_names.append(sub)
        elif current_main:
            col_names.append(current_main)
        else:
            col_names.append(f"_col{idx}_")

    col_names[0] = "_question_"
    col_names[1] = "_answer_"

    # De-duplicate data column names
    seen_names: dict[str, int] = {}
    for i, name in enumerate(col_names):
        if i < 2:
            continue
        if name in seen_names:
            seen_names[name] += 1
            col_names[i] = f"{name}_{seen_names[name]}"
        else:
            seen_names[name] = 0

    raw.columns = col_names
    data = raw.iloc[2:].reset_index(drop=True)

    data_cols = [c for c in col_names if c not in ("_question_", "_answer_")]

    questions: OrderedDict = OrderedDict()
    seen_questions: set = set()

    current_q: str | None = None
    current_rows: list[dict] = []

    def _flush():
        nonlocal current_q, current_rows
        if current_q is None or not current_rows:
            current_q = None
            current_rows = []
            return

        q_key = current_q.strip()

        if q_key.lower().strip() in SKIP_QUESTION_EXACT:
            current_q = None
            current_rows = []
            return

        if q_key in seen_questions:
            current_q = None
            current_rows = []
            return

        seen_questions.add(q_key)
        df_block = pd.DataFrame(current_rows)

        # Force _answer_ to clean strings
        df_block["_answer_"] = (
            df_block["_answer_"]
            .fillna("")
            .astype(str)
            .apply(_clean_answer)
        )

        # ── Extract n= value BEFORE dropping utility rows ──
        n_value = "?"
        for idx_row in range(len(df_block)):
            ans = df_block.iloc[idx_row]["_answer_"]
            if _row_has_n(ans):
                # Try first available data column for the n value
                for dc in data_cols:
                    if dc not in df_block.columns:
                        continue
                    raw_n = df_block.iloc[idx_row][dc]
                    if _is_empty(raw_n):
                        continue
                    n_str = (
                        str(raw_n)
                        .replace("%", "")
                        .replace(",", "")
                        .replace("\xa0", "")
                        .strip()
                    )
                    try:
                        n_num = int(float(n_str))
                        if n_num > 0:
                            n_value = str(n_num)
                            break
                    except (ValueError, TypeError):
                        continue
                if n_value != "?":
                    break

        # ── Drop utility rows (n, %, topbox, bottombox) ──
        keep_mask = ~df_block["_answer_"].apply(_is_utility_row)
        df_clean = df_block[keep_mask].copy()

        # ── Convert data columns to numeric percentages ──
        for col in data_cols:
            if col not in df_clean.columns:
                continue

            # Clean string values
            series = (
                df_clean[col]
                .fillna("")
                .astype(str)
                .str.replace("%", "", regex=False)
                .str.replace(",", ".", regex=False)
                .str.replace("\xa0", "", regex=False)
                .str.strip()
            )
            df_clean[col] = pd.to_numeric(series, errors="coerce")

        # If all values are in 0-1 range (decimals), scale to 0-100
        all_vals = []
        for col in data_cols:
            if col in df_clean.columns:
                all_vals.extend(df_clean[col].dropna().tolist())
        if all_vals and max(all_vals) <= 1.0:
            for col in data_cols:
                if col in df_clean.columns:
                    df_clean[col] = df_clean[col] * 100

        answer_opts = df_clean["_answer_"].dropna().str.strip().tolist()
        # Extra filter: remove any answer that is empty after cleaning
        answer_opts = [a for a in answer_opts if a]

        questions[q_key] = {
            "df": df_clean.reset_index(drop=True),
            "n_value": n_value,
            "answer_options": answer_opts,
        }

        current_q = None
        current_rows = []

    for _, row in data.iterrows():
        q_cell = str(row["_question_"]).strip() if not _is_empty(row["_question_"]) else ""
        a_cell = str(row["_answer_"]).strip() if not _is_empty(row["_answer_"]) else ""

        # Skip footnote rows (these ARE block separators)
        if _should_skip_row(q_cell, a_cell):
            _flush()
            current_q = None
            continue

        # Empty row = end of block
        if not q_cell and not a_cell:
            _flush()
            continue

        # New question starts
        if q_cell:
            _flush()
            current_q = q_cell

        # Add row to current block (including n/% — they'll be filtered later)
        if current_q is not None and a_cell:
            current_rows.append(row.to_dict())

    _flush()

    return questions, data_cols


# =====================================================================
# 2. AUTO-DETECT CHART TYPE
# =====================================================================
def detect_chart_type(answer_options: list[str]) -> str:
    for opt in answer_options:
        low = opt.lower().strip()
        if any(kw in low for kw in STACKED_KEYWORDS):
            return "100% Gestapeld horizontaal"
        if any(low.startswith(pf) for pf in STACKED_PREFIXES):
            return "100% Gestapeld horizontaal"
    return "Staafdiagram"


# =====================================================================
# 3. POWERPOINT GENERATION
# =====================================================================

def _sort_bar_df(df: pd.DataFrame, data_cols: list[str]) -> pd.DataFrame:
    if not data_cols:
        return df

    first_col = data_cols[0]
    answers_lower = df["_answer_"].str.strip().str.lower()

    bottom_mask = answers_lower.isin(BOTTOM_LABELS)
    penult_mask = answers_lower.apply(
        lambda x: any(x.startswith(p) or x == p for p in PENULTIMATE_LABELS)
    )
    normal_mask = ~(bottom_mask | penult_mask)

    df_normal = df[normal_mask].sort_values(by=first_col, ascending=True)
    df_penult = df[penult_mask]
    df_bottom = df[bottom_mask]

    return pd.concat([df_bottom, df_penult, df_normal], ignore_index=True)


def _set_chart_title(chart, question: str, n_value: str, group_id: str):
    """Add question + basis as chart title (two paragraphs)."""
    chart.has_title = True
    tf = chart.chart_title.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = question
    _set_font(p.font, size=Pt(10), bold=True, color=DARK_GREY)
    p.alignment = PP_ALIGN.CENTER

    basis_text = f"Basis: {group_id} (n={n_value})" if group_id else f"Basis: totaal (n={n_value})"
    p2 = tf.add_paragraph()
    p2.text = basis_text
    _set_font(p2.font, size=Pt(10), bold=False, color=MID_GREY)
    p2.alignment = PP_ALIGN.CENTER


def _clean_axes(chart):
    for axis_attr in ("value_axis", "category_axis"):
        ax = getattr(chart, axis_attr, None)
        if ax is None:
            continue
        ax.has_major_gridlines = False
        ax.has_minor_gridlines = False
        try:
            ax.major_tick_mark = 2  # NONE
            ax.minor_tick_mark = 2
        except Exception:
            pass
        try:
            ax.format.line.fill.background()
        except Exception:
            pass

    if chart.value_axis:
        chart.value_axis.visible = False


def _build_bar_slide(prs, question: str, info: dict,
                     selected_cols: list[str], group_id: str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    df = info["df"].copy()
    n_value = info["n_value"]

    keep = ["_answer_"] + [c for c in selected_cols if c in df.columns]
    df = df[[c for c in keep if c in df.columns]]
    chart_data_cols = [c for c in selected_cols if c in df.columns]

    if df.empty or not chart_data_cols:
        return

    # Extra safety: drop any remaining utility rows
    df = df[~df["_answer_"].apply(_is_utility_row)].copy()
    if df.empty:
        return

    df = _sort_bar_df(df, chart_data_cols)

    categories = df["_answer_"].tolist()
    chart_data = CategoryChartData()
    chart_data.categories = categories
    for col in chart_data_cols:
        chart_data.add_series(col, df[col].fillna(0).tolist())

    # Position chart within slide guides (~0.5" margins)
    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.5), Inches(0.4), Inches(12.33), Inches(6.7),
        chart_data,
    )

    chart = chart_frame.chart
    _set_chart_title(chart, question, n_value, group_id)
    plot = chart.plots[0]
    plot.gap_width = 60

    _clean_axes(chart)

    if chart.category_axis:
        cat_font = chart.category_axis.tick_labels.font
        _set_font(cat_font, size=Pt(10), color=DARK_GREY)

    answers_lower = df["_answer_"].str.strip().str.lower().tolist()
    colors = [BAR_PRIMARY, BAR_SECONDARY, "#FF6B81", "#A855F7"]
    num_series = len(chart_data_cols)

    for s_idx, series in enumerate(plot.series):
        base_color = colors[s_idx % len(colors)]
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = hex_to_rgb(base_color)

        for pt_idx, ans in enumerate(answers_lower):
            if ans in BOTTOM_LABELS:
                pt = series.points[pt_idx]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = hex_to_rgb(BAR_GREY)

        # Data labels — percentage at end of each bar
        series.has_data_labels = True
        dl = series.data_labels
        dl.font.name = FONT_NAME
        dl.font.size = Pt(10)
        dl.font.color.rgb = DARK_GREY
        dl.number_format = '0"%"'
        dl.number_format_is_linked = False
        try:
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
        except Exception:
            pass

    if num_series <= 1:
        chart.has_legend = False
    else:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        _set_font(chart.legend.font, size=Pt(10))


def _build_stacked_slide(prs, question: str, info: dict,
                         selected_cols: list[str], group_id: str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    df = info["df"].copy()
    n_value = info["n_value"]

    keep = ["_answer_"] + [c for c in selected_cols if c in df.columns]
    df = df[[c for c in keep if c in df.columns]]
    chart_data_cols = [c for c in selected_cols if c in df.columns]

    if df.empty or not chart_data_cols:
        return

    # Extra safety: drop any remaining utility rows
    df = df[~df["_answer_"].apply(_is_utility_row)].copy()
    if df.empty:
        return

    chart_data = CategoryChartData()
    chart_data.categories = chart_data_cols

    answer_labels = df["_answer_"].str.strip().tolist()
    for _, row in df.iterrows():
        label = str(row["_answer_"]).strip()
        vals = [float(row[c]) if pd.notna(row[c]) else 0.0 for c in chart_data_cols]
        chart_data.add_series(label, vals)

    # Position chart within slide guides (~0.5" margins)
    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_STACKED_100,
        Inches(0.5), Inches(0.4), Inches(12.33), Inches(6.7),
        chart_data,
    )

    chart = chart_frame.chart
    _set_chart_title(chart, question, n_value, group_id)
    plot = chart.plots[0]
    plot.gap_width = 60
    plot.overlap = 100

    _clean_axes(chart)

    if chart.category_axis:
        cat_font = chart.category_axis.tick_labels.font
        _set_font(cat_font, size=Pt(10), color=DARK_GREY)

    for s_idx, series in enumerate(plot.series):
        label = answer_labels[s_idx] if s_idx < len(answer_labels) else ""
        matched = _match_stacked_color(label.lower().strip())
        color = matched if matched else BAR_PRIMARY
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = hex_to_rgb(color)

        series.has_data_labels = True
        dl = series.data_labels
        dl.font.name = FONT_NAME
        dl.font.size = Pt(10)
        dl.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        dl.number_format = '[>=4]0"%";""'
        dl.number_format_is_linked = False
        dl.position = XL_LABEL_POSITION.CENTER

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    _set_font(chart.legend.font, size=Pt(10))


def _match_stacked_color(text: str) -> str | None:
    text = text.strip().lower()
    if text in STACKED_COLOR_MAP:
        return STACKED_COLOR_MAP[text]
    for key, color in STACKED_COLOR_MAP.items():
        if key in text or text in key:
            return color
    return None


def generate_pptx(questions_data: OrderedDict, config_df: pd.DataFrame,
                  selected_cols: list[str], template_file=None) -> bytes:
    if template_file is not None:
        prs = Presentation(template_file)
    else:
        prs = Presentation()
        prs.slide_width = SLIDE_WIDTH
        prs.slide_height = SLIDE_HEIGHT

    for _, row in config_df.iterrows():
        if not row.get("Exporteren", False):
            continue

        q_text = row["Vraag"]
        chart_type = row["Grafiektype"]
        group_id = str(row.get("Groep_ID", "")).strip()

        if q_text not in questions_data:
            continue

        info = questions_data[q_text]

        if chart_type == "100% Gestapeld horizontaal":
            _build_stacked_slide(prs, q_text, info, selected_cols, group_id)
        else:
            _build_bar_slide(prs, q_text, info, selected_cols, group_id)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()


# =====================================================================
# 4. STREAMLIT APP
# =====================================================================
def main():
    st.set_page_config(page_title="SPSS -> PowerPoint", layout="wide")
    st.title("SPSS Excel -> PowerPoint Rapportage")

    with st.sidebar:
        st.header("1. Upload Excel")
        uploaded = st.file_uploader("Kies een .xlsx bestand", type=["xlsx"])

        st.header("Template (optioneel)")
        template = st.file_uploader(
            "Upload een .pptx template",
            type=["pptx"],
            help="Optioneel: upload een PowerPoint-template met jouw huisstijl en hulplijnen.",
        )
        if template:
            st.session_state.template_file = io.BytesIO(template.read())
        elif "template_file" not in st.session_state:
            st.session_state.template_file = None

        if uploaded:
            if ("uploaded_name" not in st.session_state
                    or st.session_state.uploaded_name != uploaded.name):
                with st.spinner("Bezig met parsen van SPSS data..."):
                    try:
                        questions, data_cols = parse_spss_excel(uploaded)
                    except Exception as e:
                        st.error(f"Fout bij het inlezen: {e}")
                        import traceback
                        st.code(traceback.format_exc())
                        return

                st.session_state.questions = questions
                st.session_state.uploaded_name = uploaded.name
                st.session_state.available_cols = data_cols

                config_rows = []
                for q_text, info in questions.items():
                    ctype = detect_chart_type(info["answer_options"])
                    config_rows.append({
                        "Exporteren": True,
                        "Vraag": q_text,
                        "Grafiektype": ctype,
                        "Basis (n=)": f"n={info['n_value']}",
                        "Groep_ID": "",
                    })
                st.session_state.config_df = pd.DataFrame(config_rows)
                st.success(f"{len(questions)} vragen gevonden!")

        if "available_cols" in st.session_state and st.session_state.available_cols:
            st.header("2. Kolommen")
            selected = st.multiselect(
                "Selecteer uitsplitsingen",
                options=st.session_state.available_cols,
                default=st.session_state.available_cols[:1],
                help="Kies welke kolommen (Totaal, Man, Vrouw, ...) je in de grafieken wilt.",
            )
            st.session_state.selected_cols = selected

    if "config_df" not in st.session_state:
        st.info("Upload een SPSS Excel-bestand via de sidebar om te beginnen.")
        return

    st.header("Regie Tabel")
    st.caption("Pas instellingen per vraag aan. Vink 'Exporteren' aan voor opname in de PowerPoint.")

    edited = st.data_editor(
        st.session_state.config_df,
        column_config={
            "Exporteren": st.column_config.CheckboxColumn("Exporteren", default=True),
            "Vraag": st.column_config.TextColumn("Vraag", disabled=True, width="large"),
            "Grafiektype": st.column_config.SelectboxColumn(
                "Grafiektype",
                options=["Staafdiagram", "100% Gestapeld horizontaal"],
                width="medium",
            ),
            "Basis (n=)": st.column_config.TextColumn("Basis (n=)", disabled=True, width="small"),
            "Groep_ID": st.column_config.TextColumn("Groep_ID", width="small"),
        },
        use_container_width=True,
        num_rows="fixed",
        hide_index=True,
        key="regie_editor",
    )

    st.session_state.config_df = edited

    # ── Debug info (collapsed) ──
    with st.expander("Debug: data per vraag"):
        questions = st.session_state.questions
        q_choice = st.selectbox("Kies een vraag", list(questions.keys()), key="debug_q")
        if q_choice:
            info = questions[q_choice]
            st.write(f"**n = {info['n_value']}**")
            st.write(f"**Antwoordopties ({len(info['answer_options'])}):** {info['answer_options']}")
            display_df = info["df"].drop(columns=["_question_"], errors="ignore")
            st.dataframe(display_df, use_container_width=True, hide_index=True)

    st.divider()
    selected_cols = st.session_state.get("selected_cols", [])

    if not selected_cols:
        st.warning("Selecteer minstens een kolom in de sidebar.")
        return

    export_count = int(edited["Exporteren"].sum())

    if st.button(
        f"Genereer PowerPoint ({export_count} slides)",
        type="primary",
        disabled=(export_count == 0),
    ):
        with st.spinner("PowerPoint wordt gegenereerd..."):
            try:
                tpl = st.session_state.get("template_file")
                if tpl:
                    tpl.seek(0)
                pptx_bytes = generate_pptx(
                    st.session_state.questions, edited, selected_cols,
                    template_file=tpl,
                )
            except Exception as e:
                st.error(f"Fout bij genereren: {e}")
                import traceback
                st.code(traceback.format_exc())
                return

        st.success("PowerPoint is klaar!")
        st.download_button(
            label="Download rapportage.pptx",
            data=pptx_bytes,
            file_name="rapportage.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )


if __name__ == "__main__":
    main()
