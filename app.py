import streamlit as st
import pandas as pd
from bot import fill_form_from_excel
import io

df = pd.read_excel("data/template.xlsx")

st.title("EMIS Form Filler")

output = io.BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    df.to_excel(writer, index=False)
excel_data = output.getvalue()

st.download_button(
    label="Download template",
    data=excel_data,
    file_name="template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    icon=":material/download:",
)

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    data = pd.read_excel(uploaded_file)
    st.write(data)
    
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Run"):
        try:
            fill_form_from_excel(data, username, password)
            st.success("Form filling process completed successfully!")
        except Exception as e:
            st.error(f"An error occurred: {e}")