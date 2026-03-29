import io
import sys
from datetime import datetime

import pandas as pd


def main():
    import streamlit as st

    st.set_page_config(
        page_title="DataMatch Pro",
        page_icon= "logo.png",
        layout="wide"
    )
    st.image("logo.png", width=250)

    def read_excel_with_sheet_selector(uploaded_file, label):
    import streamlit as st
    import pandas as pd

    # Get all sheet names
    excel_file = pd.ExcelFile(uploaded_file)
    sheets = excel_file.sheet_names

    # Let user select sheet
    selected_sheet = st.selectbox(
        f"Select sheet for {label}",
        sheets,
        key=f"{label}_sheet"
    )

    # Read selected sheet
   def read_excel_with_sheet_selector(uploaded_file, label):
    import streamlit as st
    import pandas as pd

    # Get all sheet names
    excel_file = pd.ExcelFile(uploaded_file)
    sheets = excel_file.sheet_names

    # Let user select sheet
    selected_sheet = st.selectbox(
        f"Select sheet for {label}",
        sheets,
        key=f"{label}_sheet"
    )

    # Read selected sheet
    df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
    df.columns = df.columns.astype(str).str.strip()

    return df

    def clean_for_match(series: pd.Series) -> pd.Series:
        return (
            series.astype(str)
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

    st.title("DataMatch Pro")
    st.caption("Match files, bring selected columns, or extract common data into a new Excel file.")

    st.sidebar.header("Settings")

    mode = st.sidebar.radio(
        "Select mode",
        options=[
            "Match Sheet 1 with Sheet 2 and bring selected columns",
            "Take common data between 2 or more files"
        ]
    )

    preview_rows = st.sidebar.slider("Preview rows", min_value=5, max_value=50, value=20, step=5)

    if mode == "Match Sheet 1 with Sheet 2 and bring selected columns":
        st.subheader("Mode 1: Match Sheet 1 with Sheet 2")

        col_a, col_b = st.columns(2)
        with col_a:
            file1 = st.file_uploader("Upload Sheet 1", type=["xlsx", "xls"], key="mode1_file1")
        with col_b:
            file2 = st.file_uploader("Upload Sheet 2", type=["xlsx", "xls"], key="mode1_file2")

        if file1 and file2:
            try:
                df1 = read_excel_file(file1)
                df2 = read_excel_file(file2)

                st.success("Files loaded successfully.")

                left1, left2, left3 = st.columns(3)

                with left1:
                    match_cols_1 = st.multiselect(
                        "Match columns from Sheet 1",
                        options=list(df1.columns)
                    )

                with left2:
                    match_cols_2 = st.multiselect(
                        "Match columns from Sheet 2",
                        options=list(df2.columns)
                    )

                with left3:
                    bring_cols = st.multiselect(
                        "Columns to bring from Sheet 2",
                        options=list(df2.columns)
                    )

                keep_all_sheet1 = st.checkbox(
                    "Keep all rows from Sheet 1 even if not matched",
                    value=True
                )

                export_only_selected = st.checkbox(
                    "Export only selected columns",
                    value=False
                )

                final_output_cols = []
                if export_only_selected:
                    possible_result_cols = list(dict.fromkeys(
                        list(df1.columns) + list(df2.columns) + [f"{c}_from_sheet2" for c in df2.columns]
                    ))
                    final_output_cols = st.multiselect(
                        "Final output columns",
                        options=possible_result_cols,
                        default=list(df1.columns)
                    )

                if st.button("Build Result", type="primary", key="mode1_build"):
                    if not match_cols_1 or not match_cols_2:
                        st.error("Please select matching columns from both sheets.")
                    elif len(match_cols_1) != len(match_cols_2):
                        st.error("The number of matching columns must be the same in both sheets.")
                    elif not bring_cols:
                        st.error("Please select at least one column to bring from Sheet 2.")
                    else:
                        df1_work = df1.copy()
                        df2_work = df2.copy()

                        key_cols_1 = []
                        key_cols_2 = []

                        for i, (col1, col2) in enumerate(zip(match_cols_1, match_cols_2), start=1):
                            key1 = f"__match1_{i}__"
                            key2 = f"__match2_{i}__"
                            df1_work[key1] = clean_for_match(df1_work[col1])
                            df2_work[key2] = clean_for_match(df2_work[col2])
                            key_cols_1.append(key1)
                            key_cols_2.append(key2)

                        df2_subset = df2_work[key_cols_2 + bring_cols].copy()

                        rename_keys = {k2: k1 for k1, k2 in zip(key_cols_1, key_cols_2)}
                        df2_subset = df2_subset.rename(columns=rename_keys)

                        rename_bring = {}
                        for col in bring_cols:
                            if col in df1_work.columns:
                                rename_bring[col] = f"{col}_from_sheet2"
                        df2_subset = df2_subset.rename(columns=rename_bring)

                        merge_type = "left" if keep_all_sheet1 else "inner"

                        result = pd.merge(
                            df1_work,
                            df2_subset,
                            on=key_cols_1,
                            how=merge_type
                        )

                        result.drop(columns=key_cols_1, inplace=True, errors="ignore")

                        if export_only_selected:
                            missing_cols = [c for c in final_output_cols if c not in result.columns]
                            for c in missing_cols:
                                result[c] = ""
                            result = result[final_output_cols]

                        st.session_state["result_df"] = result
                        st.session_state["result_name"] = f"matched_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

                if "result_df" in st.session_state:
                    result = st.session_state["result_df"]

                    st.markdown("### Result Preview")
                    st.dataframe(safe_preview(result, preview_rows), use_container_width=True)

                    st.info(f"Rows in result: {len(result):,}")

                    excel_data = to_excel_bytes(result)
                    st.download_button(
                        label="Download Excel",
                        data=excel_data,
                        file_name=st.session_state.get("result_name", "matched_output.xlsx"),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"Failed to process files: {e}")

    else:
        st.subheader("Mode 2: Take common data between 2 or more files")

        files = st.file_uploader(
            "Upload 2 or more Excel files",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="mode2_files"
        )

        if files and len(files) >= 2:
            try:
                dfs = []
                for i, f in enumerate(files):
                    df = read_excel_with_sheet_selector(f, f"File {i+1}")
                    dfs.append(df)

                common_cols = set(dfs[0].columns)
                union_cols = set(dfs[0].columns)

                for df in dfs[1:]:
                    common_cols &= set(df.columns)
                    union_cols |= set(df.columns)

                common_cols = sorted(list(common_cols))
                union_cols = sorted(list(union_cols))

                st.success(f"{len(files)} files loaded successfully.")

                col1, col2 = st.columns(2)

                with col1:
                    common_match_cols = st.multiselect(
                        "Common columns used to identify the same rows across all files",
                        options=common_cols
                    )

                with col2:
                    output_cols = st.multiselect(
                        "Columns to export in the new Excel",
                        options=union_cols
                    )

                ignore_missing_output_cols = st.checkbox(
                    "Ignore selected output columns that do not exist in some files",
                    value=False
                )

                keep_duplicates = st.checkbox(
                    "Keep duplicate rows in final result",
                    value=False
                )

                if st.button("Build Result", type="primary", key="mode2_build"):
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

                            st.session_state["result_df"] = result
                            st.session_state["result_name"] = f"common_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

                if "result_df" in st.session_state:
                    result = st.session_state["result_df"]

                    st.markdown("### Result Preview")
                    st.dataframe(safe_preview(result, preview_rows), use_container_width=True)

                    st.info(f"Rows in result: {len(result):,}")

                    excel_data = to_excel_bytes(result)
                    st.download_button(
                        label="Download Excel",
                        data=excel_data,
                        file_name=st.session_state.get("result_name", "common_output.xlsx"),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"Failed to process files: {e}")

        elif files and len(files) < 2:
            st.warning("Please upload at least 2 files.")


main()
