import io
import json
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process


# =========================
# PAGE CONFIG
# =========================
st.set_page_config(
    page_title="DataMatch Pro",
    page_icon="logo.png",
    layout="wide"
)


# =========================
# UI HELPERS
# =========================
def show_logo():
    logo_path = Path("logo.png")
    if logo_path.exists():
        col1, col2 = st.columns([1, 5])
        with col1:
            st.image(str(logo_path), width=100)
        with col2:
            st.title("DataMatch Pro")
            st.caption("Fast Excel matching, smart matching, analytics, and reusable client mappings.")
    else:
        st.title("DataMatch Pro")
        st.caption("Fast Excel matching, smart matching, analytics, and reusable client mappings.")


# =========================
# FILE / DATA HELPERS
# =========================
@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes):
    excel = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    return excel.sheet_names


@st.cache_data(show_spinner=False)
def read_excel_sheet(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, engine="openpyxl")
    df.columns = df.columns.astype(str).str.strip()
    return df


def read_excel_with_sheet_selector(file_bytes: bytes, label: str, key_prefix: str, default_sheet: str | None = None):
    sheets = get_sheet_names(file_bytes)

    if len(sheets) == 1:
        selected_sheet = sheets[0]
        st.info(f"{label}: only one sheet found → **{selected_sheet}**")
    else:
        default_index = sheets.index(default_sheet) if default_sheet in sheets else 0
        selected_sheet = st.selectbox(
            f"Select sheet for {label}",
            sheets,
            index=default_index,
            key=f"{key_prefix}_sheet"
        )

    df = read_excel_sheet(file_bytes, selected_sheet)
    return df, selected_sheet


def clean_for_match(series: pd.Series) -> pd.Series:
    return (
        series.astype(str)
        .fillna("")
        .str.lower()
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Output")
    output.seek(0)
    return output.getvalue()


def safe_preview(df: pd.DataFrame, n: int = 20) -> pd.DataFrame:
    return df.head(n).copy()


def parse_template(uploaded_json):
    if uploaded_json is None:
        return {}
    try:
        content = uploaded_json.read().decode("utf-8")
        return json.loads(content)
    except Exception:
        return {}


def filter_defaults(options, defaults):
    return [x for x in defaults if x in options]


def clear_result(mode_name: str):
    st.session_state.pop(f"result_df_{mode_name}", None)
    st.session_state.pop(f"result_name_{mode_name}", None)


def show_missing_values_chart(df: pd.DataFrame, title: str):
    missing = df.isna().sum()
    missing = missing[missing > 0].sort_values(ascending=False)
    if not missing.empty:
        st.markdown(f"#### {title}")
        st.bar_chart(missing)


# =========================
# SMART MATCHING
# =========================
def build_fuzzy_key_map(source_keys, target_keys, threshold):
    target_unique = pd.Series(target_keys).dropna().astype(str).unique().tolist()
    mapping = {}
    scores = {}

    for val in pd.Series(source_keys).dropna().astype(str).unique():
        if val == "":
            mapping[val] = None
            scores[val] = 0
            continue

        best = process.extractOne(val, target_unique, scorer=fuzz.token_sort_ratio)
        if best:
            best_value, best_score, _ = best
            if best_score >= threshold:
                mapping[val] = best_value
                scores[val] = best_score
            else:
                mapping[val] = None
                scores[val] = best_score
        else:
            mapping[val] = None
            scores[val] = 0

    return mapping, scores


# =========================
# HEADER
# =========================
show_logo()


# =========================
# SIDEBAR
# =========================
st.sidebar.header("Settings")

mode = st.sidebar.radio(
    "Select mode",
    options=[
        "Mode 1 - Match Sheet 1 with Sheet 2 and bring selected columns",
        "Mode 2 - Take common data between 2 or more files"
    ]
)

preview_rows = st.sidebar.slider("Preview rows", min_value=5, max_value=50, value=15, step=5)

st.sidebar.markdown("### Mapping Template")
uploaded_template = st.sidebar.file_uploader("Load saved mapping (.json)", type=["json"])
loaded_template = parse_template(uploaded_template)

template_name = st.sidebar.text_input(
    "Template name",
    value=loaded_template.get("template_name", "client_mapping")
)

st.sidebar.caption(
    "Templates are downloaded as JSON and can be uploaded later. "
    "For permanent team-wide saved templates, add a database later."
)


# =========================
# MODE 1
# =========================
if mode == "Mode 1 - Match Sheet 1 with Sheet 2 and bring selected columns":
    mode_name = "mode1"
    st.subheader("Mode 1: Match Sheet 1 with Sheet 2")

    st.markdown(
        """
        **Use this mode when you want to:**
        - match two Excel files
        - bring selected columns from File 2 into File 1
        - optionally use smart matching for names
        """
    )

    file_col1, file_col2 = st.columns(2)
    with file_col1:
        file1 = st.file_uploader("Upload File 1", type=["xlsx", "xls"], key="mode1_file1")
    with file_col2:
        file2 = st.file_uploader("Upload File 2", type=["xlsx", "xls"], key="mode1_file2")

    if file1 and file2:
        try:
            file1_bytes = file1.getvalue()
            file2_bytes = file2.getvalue()

            default_sheet1 = loaded_template.get("selected_sheet1")
            default_sheet2 = loaded_template.get("selected_sheet2")

            df1, selected_sheet1 = read_excel_with_sheet_selector(
                file1_bytes, "File 1", "mode1_file1", default_sheet1
            )
            df2, selected_sheet2 = read_excel_with_sheet_selector(
                file2_bytes, "File 2", "mode1_file2", default_sheet2
            )

            st.success("Files loaded successfully.")

            prev1, prev2 = st.columns(2)
            with prev1:
                st.markdown(f"### Preview - File 1 ({selected_sheet1})")
                st.dataframe(safe_preview(df1, preview_rows), use_container_width=True)
            with prev2:
                st.markdown(f"### Preview - File 2 ({selected_sheet2})")
                st.dataframe(safe_preview(df2, preview_rows), use_container_width=True)

            st.markdown("### Column Selection")
            col1, col2, col3 = st.columns(3)

            default_match_cols_1 = filter_defaults(list(df1.columns), loaded_template.get("match_cols_1", []))
            default_match_cols_2 = filter_defaults(list(df2.columns), loaded_template.get("match_cols_2", []))
            default_bring_cols = filter_defaults(list(df2.columns), loaded_template.get("bring_cols", []))

            with col1:
                match_cols_1 = st.multiselect(
                    "Match columns from File 1",
                    options=list(df1.columns),
                    default=default_match_cols_1
                )

            with col2:
                match_cols_2 = st.multiselect(
                    "Match columns from File 2",
                    options=list(df2.columns),
                    default=default_match_cols_2
                )

            with col3:
                bring_cols = st.multiselect(
                    "Columns to bring from File 2",
                    options=list(df2.columns),
                    default=default_bring_cols
                )

            st.markdown("### Matching Options")
            opt1, opt2, opt3 = st.columns(3)

            with opt1:
                keep_all_sheet1 = st.checkbox(
                    "Keep all rows from File 1",
                    value=loaded_template.get("keep_all_sheet1", True)
                )

            with opt2:
                use_smart_matching = st.checkbox(
                    "Use smart matching (fuzzy name similarity)",
                    value=loaded_template.get("use_smart_matching", False),
                    help="Best used when matching one text column such as Employee Name."
                )

            with opt3:
                fuzzy_threshold = st.slider(
                    "Smart matching threshold",
                    min_value=50,
                    max_value=100,
                    value=int(loaded_template.get("fuzzy_threshold", 85)),
                    step=1,
                    disabled=not use_smart_matching
                )

            st.markdown("### Output Options")
            export_only_selected = st.checkbox(
                "Export only selected columns",
                value=loaded_template.get("export_only_selected", False)
            )

            possible_result_cols = list(dict.fromkeys(
                list(df1.columns)
                + list(df2.columns)
                + [f"{c}_from_file2" for c in df2.columns]
                + ["match_score", "match_status"]
            ))

            default_final_cols = filter_defaults(
                possible_result_cols,
                loaded_template.get("final_output_cols", list(df1.columns))
            )

            final_output_cols = []
            if export_only_selected:
                final_output_cols = st.multiselect(
                    "Final output columns",
                    options=possible_result_cols,
                    default=default_final_cols
                )

            # Current mapping template
            current_template = {
                "template_name": template_name,
                "mode": mode_name,
                "selected_sheet1": selected_sheet1,
                "selected_sheet2": selected_sheet2,
                "match_cols_1": match_cols_1,
                "match_cols_2": match_cols_2,
                "bring_cols": bring_cols,
                "keep_all_sheet1": keep_all_sheet1,
                "use_smart_matching": use_smart_matching,
                "fuzzy_threshold": fuzzy_threshold,
                "export_only_selected": export_only_selected,
                "final_output_cols": final_output_cols,
            }

            st.sidebar.download_button(
                "Download current mapping",
                data=json.dumps(current_template, indent=2),
                file_name=f"{template_name or 'client_mapping'}_mode1.json",
                mime="application/json"
            )

            if st.button("Build Result", type="primary", key="mode1_build"):
                clear_result(mode_name)

                if not match_cols_1 or not match_cols_2:
                    st.error("Please select matching columns from both files.")
                elif len(match_cols_1) != len(match_cols_2):
                    st.error("The number of matching columns must be the same in both files.")
                elif not bring_cols:
                    st.error("Please select at least one column to bring from File 2.")
                else:
                    df1_work = df1.copy()
                    df2_work = df2.copy()

                    # SMART MATCHING
                    if use_smart_matching and len(match_cols_1) == 1 and len(match_cols_2) == 1:
                        key1 = match_cols_1[0]
                        key2 = match_cols_2[0]

                        df1_work["__match1__"] = clean_for_match(df1_work[key1])
                        df2_work["__match2__"] = clean_for_match(df2_work[key2])

                        key_map, score_map = build_fuzzy_key_map(
                            df1_work["__match1__"],
                            df2_work["__match2__"],
                            fuzzy_threshold
                        )

                        df1_work["__mapped_key__"] = df1_work["__match1__"].map(key_map)
                        df1_work["match_score"] = df1_work["__match1__"].map(score_map).fillna(0)

                        df2_subset = df2_work[["__match2__"] + bring_cols].copy()

                        rename_bring = {}
                        for col in bring_cols:
                            if col in df1_work.columns:
                                rename_bring[col] = f"{col}_from_file2"
                        df2_subset = df2_subset.rename(columns=rename_bring)

                        merge_type = "left" if keep_all_sheet1 else "inner"

                        result = pd.merge(
                            df1_work,
                            df2_subset,
                            left_on="__mapped_key__",
                            right_on="__match2__",
                            how=merge_type
                        )

                        result["match_status"] = result["__match2__"].notna().map(
                            lambda x: "Matched" if x else "Unmatched"
                        )

                        result.drop(columns=["__match1__", "__mapped_key__", "__match2__"], inplace=True, errors="ignore")

                    # EXACT MATCHING
                    else:
                        key_cols_1 = []
                        key_cols_2 = []

                        for i, (col1_name, col2_name) in enumerate(zip(match_cols_1, match_cols_2), start=1):
                            key1 = f"__match1_{i}__"
                            key2 = f"__match2_{i}__"
                            df1_work[key1] = clean_for_match(df1_work[col1_name])
                            df2_work[key2] = clean_for_match(df2_work[col2_name])
                            key_cols_1.append(key1)
                            key_cols_2.append(key2)

                        df2_subset = df2_work[key_cols_2 + bring_cols].copy()

                        rename_keys = {k2: k1 for k1, k2 in zip(key_cols_1, key_cols_2)}
                        df2_subset = df2_subset.rename(columns=rename_keys)

                        rename_bring = {}
                        for col in bring_cols:
                            if col in df1_work.columns:
                                rename_bring[col] = f"{col}_from_file2"
                        df2_subset = df2_subset.rename(columns=rename_bring)

                        merge_type = "left" if keep_all_sheet1 else "inner"

                        result = pd.merge(
                            df1_work,
                            df2_subset,
                            on=key_cols_1,
                            how=merge_type
                        )

                        brought_final_cols = [rename_bring.get(c, c) for c in bring_cols]
                        result["match_status"] = result[brought_final_cols].notna().any(axis=1).map(
                            lambda x: "Matched" if x else "Unmatched"
                        )
                        result["match_score"] = None

                        result.drop(columns=key_cols_1, inplace=True, errors="ignore")

                    if export_only_selected:
                        missing_cols = [c for c in final_output_cols if c not in result.columns]
                        for c in missing_cols:
                            result[c] = ""
                        result = result[final_output_cols]

                    st.session_state[f"result_df_{mode_name}"] = result
                    st.session_state[f"result_name_{mode_name}"] = (
                        f"matched_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    )

            result_key = f"result_df_{mode_name}"
            result_name_key = f"result_name_{mode_name}"

            if result_key in st.session_state:
                result = st.session_state[result_key]

                st.markdown("## Result Preview")
                st.dataframe(safe_preview(result, preview_rows), use_container_width=True)

                st.markdown("## Analytics Dashboard")
                matched_count = int((result["match_status"] == "Matched").sum()) if "match_status" in result.columns else 0
                unmatched_count = int((result["match_status"] == "Unmatched").sum()) if "match_status" in result.columns else 0
                total_rows = len(result)
                match_rate = round((matched_count / total_rows) * 100, 2) if total_rows else 0

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Total Rows", f"{total_rows:,}")
                m2.metric("Matched Rows", f"{matched_count:,}")
                m3.metric("Unmatched Rows", f"{unmatched_count:,}")
                m4.metric("Match Rate", f"{match_rate}%")

                chart_df = pd.DataFrame({
                    "Status": ["Matched", "Unmatched"],
                    "Count": [matched_count, unmatched_count]
                }).set_index("Status")
                st.bar_chart(chart_df)

                if "match_score" in result.columns and result["match_score"].notna().any():
                    st.markdown("#### Match Score Distribution")
                    score_counts = pd.cut(
                        result["match_score"].fillna(0),
                        bins=[0, 60, 70, 80, 90, 100],
                        include_lowest=True
                    ).value_counts().sort_index()
                    st.bar_chart(score_counts)

                show_missing_values_chart(result, "Missing Values by Column")

                excel_data = to_excel_bytes(result)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=st.session_state.get(result_name_key, "matched_output.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Failed to process files: {e}")


# =========================
# MODE 2
# =========================
else:
    mode_name = "mode2"
    st.subheader("Mode 2: Take common data between 2 or more files")

    st.markdown(
        """
        **Use this mode when you want to:**
        - upload multiple Excel files
        - identify common rows across all files
        - export only the columns you want
        """
    )

    files = st.file_uploader(
        "Upload 2 or more Excel files",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="mode2_files"
    )

    if files and len(files) >= 2:
        try:
            dfs = []
            selected_sheets = []

            st.markdown("### Sheet Selection")
            template_sheets = loaded_template.get("selected_sheets", [])

            for i, file in enumerate(files):
                st.markdown(f"**File {i+1}: {file.name}**")
                default_sheet = template_sheets[i] if i < len(template_sheets) else None
                df, selected_sheet = read_excel_with_sheet_selector(
                    file.getvalue(),
                    f"File {i+1}",
                    f"mode2_file_{i+1}",
                    default_sheet
                )
                dfs.append(df)
                selected_sheets.append(selected_sheet)

            st.success(f"{len(files)} files loaded successfully.")

            with st.expander("Preview Uploaded Files", expanded=False):
                for i, df in enumerate(dfs):
                    st.markdown(f"**File {i+1} - Sheet: {selected_sheets[i]}**")
                    st.dataframe(safe_preview(df, min(preview_rows, 10)), use_container_width=True)

            common_cols = set(dfs[0].columns)
            union_cols = set(dfs[0].columns)

            for df in dfs[1:]:
                common_cols &= set(df.columns)
                union_cols |= set(df.columns)

            common_cols = sorted(list(common_cols))
            union_cols = sorted(list(union_cols))

            default_common_match = filter_defaults(common_cols, loaded_template.get("common_match_cols", []))
            default_output_cols = filter_defaults(union_cols, loaded_template.get("output_cols", []))

            st.markdown("### Column Selection")
            c1, c2 = st.columns(2)

            with c1:
                common_match_cols = st.multiselect(
                    "Common columns used to identify the same rows across all files",
                    options=common_cols,
                    default=default_common_match
                )

            with c2:
                output_cols = st.multiselect(
                    "Columns to export in the new Excel",
                    options=union_cols,
                    default=default_output_cols
                )

            ignore_missing_output_cols = st.checkbox(
                "Ignore selected output columns that do not exist in some files",
                value=loaded_template.get("ignore_missing_output_cols", False)
            )

            keep_duplicates = st.checkbox(
                "Keep duplicate rows in final result",
                value=loaded_template.get("keep_duplicates", False)
            )

            current_template = {
                "template_name": template_name,
                "mode": mode_name,
                "selected_sheets": selected_sheets,
                "common_match_cols": common_match_cols,
                "output_cols": output_cols,
                "ignore_missing_output_cols": ignore_missing_output_cols,
                "keep_duplicates": keep_duplicates,
            }

            st.sidebar.download_button(
                "Download current mapping",
                data=json.dumps(current_template, indent=2),
                file_name=f"{template_name or 'client_mapping'}_mode2.json",
                mime="application/json"
            )

            if st.button("Build Result", type="primary", key="mode2_build"):
                clear_result(mode_name)

                if not common_match_cols:
                    st.error("Please select the common columns used for matching.")
                elif not output_cols:
                    st.error("Please select at least one output column.")
                else:
                    processed = []
                    common_key_names = []

                    for df in dfs:
                        df_copy = df.copy()
                        for i, col in enumerate(common_match_cols, start=1):
                            key_name = f"__common_key_{i}__"
                            if key_name not in common_key_names:
                                common_key_names.append(key_name)
                            df_copy[key_name] = clean_for_match(df_copy[col])
                        processed.append(df_copy)

                    result = processed[0].copy()

                    for next_df in processed[1:]:
                        key_only = next_df[common_key_names].drop_duplicates()
                        result = pd.merge(result, key_only, on=common_key_names, how="inner")

                    result.drop(columns=common_key_names, inplace=True, errors="ignore")

                    final_cols = []
                    for col in output_cols:
                        if col in result.columns:
                            final_cols.append(col)
                        elif not ignore_missing_output_cols:
                            result[col] = ""
                            final_cols.append(col)

                    if not final_cols:
                        st.error("None of the selected output columns are available in the final result.")
                    else:
                        result = result[final_cols]

                        if not keep_duplicates:
                            result = result.drop_duplicates()

                        st.session_state[f"result_df_{mode_name}"] = result
                        st.session_state[f"result_name_{mode_name}"] = (
                            f"common_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        )

            result_key = f"result_df_{mode_name}"
            result_name_key = f"result_name_{mode_name}"

            if result_key in st.session_state:
                result = st.session_state[result_key]

                st.markdown("## Result Preview")
                st.dataframe(safe_preview(result, preview_rows), use_container_width=True)

                st.markdown("## Analytics Dashboard")
                d1, d2, d3 = st.columns(3)
                d1.metric("Total Rows", f"{len(result):,}")
                d2.metric("Columns in Result", f"{len(result.columns):,}")
                d3.metric("Duplicates Removed", "Yes" if not keep_duplicates else "No")

                show_missing_values_chart(result, "Missing Values by Column")

                excel_data = to_excel_bytes(result)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=st.session_state.get(result_name_key, "common_output.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Failed to process files: {e}")

    elif files and len(files) < 2:
        st.warning("Please upload at least 2 files.")
