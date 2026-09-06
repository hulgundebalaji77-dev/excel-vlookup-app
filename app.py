import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Advanced Excel & CSV VLOOKUP", layout="wide", page_icon="📊")

st.title("📊 प्रगत VLOOKUP वेब ॲप (Excel & CSV)")
st.write("दोन Excel/CSV फाईल्स अपलोड करा, हवी ती शीट निवडा आणि एका क्लिकवर डेटा एकत्र करा.")

def load_data(uploaded_file, key_prefix):
    """Excel किंवा CSV फाईल वाचून डेटाफ्रेम देणारे फंक्शन"""
    if uploaded_file.name.endswith(".csv"):
        return pd.read_csv(uploaded_file), None
    else:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        selected_sheet = st.selectbox(
            f"शीट निवडा ({uploaded_file.name}):",
            options=sheet_names,
            key=f"{key_prefix}_sheet"
        )
        return pd.read_excel(uploaded_file, sheet_name=selected_sheet), selected_sheet

col1, col2 = st.columns(2)

with col1:
    st.subheader("१. Main File (मूळ फाईल)")
    main_file = st.file_uploader(
        "Main File अपलोड करा (Excel/CSV)", 
        type=["xlsx", "xls", "csv"], 
        key="main"
    )

with col2:
    st.subheader("२. Lookup File (माहिती शोधायची फाईल)")
    lookup_file = st.file_uploader(
        "Lookup File अपलोड करा (Excel/CSV)", 
        type=["xlsx", "xls", "csv"], 
        key="lookup"
    )

if main_file and lookup_file:
    try:
        df_main, main_sheet = load_data(main_file, "main")
        df_lookup, lookup_sheet = load_data(lookup_file, "lookup")

        st.markdown("---")
        st.subheader("🔍 मॅचिंग पर्याय आणि कॉलम्स")

        c1, c2, c3 = st.columns(3)

        with c1:
            main_key = st.selectbox(
                "Main File मधील Key Column:",
                options=df_main.columns,
                help="ज्या कॉलमच्या आधारे डेटा शोधायचा आहे."
            )

        with c2:
            lookup_key = st.selectbox(
                "Lookup File मधील Key Column:",
                options=df_lookup.columns,
                help="Lookup फाईलमधील मॅच होणारा कॉलम."
            )

        with c3:
            available_cols = [c for c in df_lookup.columns if c != lookup_key]
            cols_to_fetch = st.multiselect(
                "Lookup फाईलमधून कोणते कॉलम्स जोडायचे आहेत?",
                options=available_cols,
                default=available_cols[:1] if available_cols else []
            )

        col_join, col_format = st.columns(2)
        with col_join:
            how_type = st.radio(
                "मॅचिंग प्रकार (Join Type):",
                options=["left", "inner", "outer"],
                format_func=lambda x: {
                    "left": "Left Join (Main फाईलमधील सर्व डेटा राहील)",
                    "inner": "Inner Join (फक्त मॅच झालेले रेकॉर्ड्स)",
                    "outer": "Outer Join (दोन्ही फाईल्समधील सर्व रेकॉर्ड्स)"
                }[x],
                horizontal=True
            )

        with col_format:
            export_format = st.radio(
                "डाऊनलोड फॉरमॅट निवडा:",
                options=["Excel (.xlsx)", "CSV (.csv)"],
                horizontal=True
            )

        if st.button("🚀 VLOOKUP चालवा / एकत्र करा", type="primary"):
            if not cols_to_fetch:
                st.warning("कृपया Lookup फाईलमधून किमान एक कॉलम निवडा.")
            else:
                with st.spinner("डेटा प्रक्रिया सुरू आहे..."):
                    # डेटा प्रकार एकसारखा करणे (Mismatch टाळण्यासाठी string मध्ये रूपांतर)
                    df_main_proc = df_main.copy()
                    df_lookup_proc = df_lookup.copy()

                    df_main_proc[main_key] = df_main_proc[main_key].astype(str).str.strip()
                    df_lookup_proc[lookup_key] = df_lookup_proc[lookup_key].astype(str).str.strip()

                    # Lookup टेबलमधील आवश्यक कॉलम्स घेऊन डुप्लिकेट्स काढणे
                    lookup_subset = df_lookup_proc[[lookup_key] + cols_to_fetch].drop_duplicates(subset=[lookup_key])

                    # Merge प्रक्रिया
                    result_df = pd.merge(
                        df_main_proc,
                        lookup_subset,
                        left_on=main_key,
                        right_on=lookup_key,
                        how=how_type
                    )

                    if main_key != lookup_key and lookup_key in result_df.columns:
                        result_df.drop(columns=[lookup_key], inplace=True)

                    st.success("✅ VLOOKUP यशस्वीरीत्या पूर्ण झाले!")
                    st.dataframe(result_df.head(25), use_container_width=True)

                    # डाऊनलोड पर्याय
                    if export_format == "Excel (.xlsx)":
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine="openpyxl") as writer:
                            result_df.to_excel(writer, index=False, sheet_name="Result")
                        output.seek(0)
                        st.download_button(
                            label="📥 Excel फाईल डाऊनलोड करा",
                            data=output,
                            file_name="Vlookup_Merged_Data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        csv_data = result_df.to_csv(index=False).encode("utf-8")
                        st.download_button(
                            label="📥 CSV फाईल डाऊनलोड करा",
                            data=csv_data,
                            file_name="Vlookup_Merged_Data.csv",
                            mime="text/csv"
                        )

    except Exception as e:
        st.error(f"प्रक्रियेदरम्यान त्रुटी आली: {e}")
else:
    st.info("कृपया वरील दोन्ही फाईल्स अपलोड करा.")
