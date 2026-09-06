import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Match & Row Update Tool", layout="wide", page_icon="📊")

st.title("📊 Excel Data Match & Row Update Web App")
st.write("File 1 मधील डेटा शोधून File 2 च्या संबंधित row मध्ये टाका / अपडेट करा.")

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
    st.subheader("१. File 1 (ज्यामधून डेटा घ्यायचा आहे - Source)")
    file1 = st.file_uploader("File 1 अपलोड करा", type=["xlsx", "xls", "csv"], key="file1")

with col2:
    st.subheader("२. File 2 (ज्यामध्ये डेटा टाकायचा आहे - Target)")
    file2 = st.file_uploader("File 2 अपलोड करा", type=["xlsx", "xls", "csv"], key="file2")

if file1 and file2:
    try:
        df1, _ = load_data(file1, "f1")
        df2, _ = load_data(file2, "f2")

        st.markdown("---")
        st.subheader("🔍 मॅचिंग आणि अपडेट पर्याय")

        c1, c2 = st.columns(2)
        with c1:
            key1 = st.selectbox(
                "File 1 मधील मॅचिंग Key Column:",
                options=df1.columns,
                help="उदा. Account No, ID, Roll No"
            )
        with c2:
            key2 = st.selectbox(
                "File 2 मधील मॅचिंग Key Column:",
                options=df2.columns,
                help="ज्या कॉलमशी File 1 चा डेटा मॅच करायचा आहे"
            )

        col_source, col_target = st.columns(2)
        with col_source:
            source_col = st.selectbox(
                "File 1 मधील कोणता व्हॅल्यू कॉलम घ्यायचा आहे? (Data to Copy):",
                options=[col for col in df1.columns if col != key1]
            )

        with col_target:
            action_type = st.radio(
                "File 2 मध्ये व्हॅल्यू कशी टाकायची आहे?",
                options=["नवीन कॉलम तयार करा (Add as New Column)", "विद्यमान कॉलम अपडेट करा (Overwrite Existing Column)"],
                horizontal=True
            )

            if action_type == "नवीन कॉलम तयार करा (Add as New Column)":
                new_col_name = st.text_input("नवीन कॉलमचे नाव लिहा:", value=f"{source_col}_matched")
                target_col = new_col_name
            else:
                target_col = st.selectbox("File 2 मधील कोणता कॉलम अपडेट करायचा आहे?", options=df2.columns)

        if st.button("🚀 डेटा मॅच करून File 2 मध्ये टाका", type="primary"):
            with st.spinner("डेटा मॅच करून अपडेट होत आहे..."):
                # डेटा कॉपी तयार करणे
                df1_clean = df1.copy()
                df2_result = df2.copy()

                # मॅचिंग अचूक होण्यासाठी दोन्ही की स्ट्रिंग स्वरूपात आणि स्पेस काढून घेणे
                df1_clean[key1] = df1_clean[key1].astype(str).str.strip()
                df2_result[key2] = df2_result[key2].astype(str).str.strip()

                # File 1 मधील key आणि source column चे डिक्शनरी मॅपिंग तयार करणे (डुप्लिकेट्स काढून)
                df1_unique = df1_clean.drop_duplicates(subset=[key1])
                mapping_dict = dict(zip(df1_unique[key1], df1_unique[source_col]))

                # File 2 मध्ये डेटा टाकणे
                if action_type == "नवीन कॉलम तयार करा (Add as New Column)":
                    df2_result[target_col] = df2_result[key2].map(mapping_dict)
                else:
                    # विद्यमान कॉलममधील डेटा अपडेट करणे (जर मॅच झाला तरच व्हॅल्यू बदलेल, नाहीतर जुनीच राहील)
                    df2_result[target_col] = df2_result[key2].map(mapping_dict).fillna(df2_result[target_col])

                st.success(f"✅ डेटा यशस्वीरीत्या मॅच झाला आणि File 2 च्या '{target_col}' कॉलम/रो मध्ये अपडेट झाला!")

                # Preview दाखवणे
                st.subheader("📋 अपडेटेड File 2 चे पूर्वावलोकन (Preview):")
                st.dataframe(df2_result.head(25), use_container_width=True)

                # डाऊनलोड पर्याय
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df2_result.to_excel(writer, index=False, sheet_name="Updated_Data")
                output.seek(0)

                st.download_button(
                    label="📥 अपडेट झालेली File 2 डाऊनलोड करा (Excel)",
                    data=output,
                    file_name="File2_Updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"त्रुटी आली: {e}")
else:
    st.info("कृपया वरील दोन्ही Excel किंवा CSV फाईल्स अपलोड करा.")
