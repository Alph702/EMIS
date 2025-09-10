import streamlit as st
import pandas as pd
from bot import fill_form_from_excel
import io
import os

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

    photos = st.file_uploader(
        "Upload images", accept_multiple_files=True, type=["jpg", "png"]
    )

    if photos:
        save_dir = "Photos"
        os.makedirs(save_dir, exist_ok=True)
        for photo in photos:
            save_path = os.path.join(save_dir, photo.name)
            with open(save_path, "wb") as f:
                f.write(photo.getbuffer())
        st.success(f"Saved {len(photos)} file(s) to: {save_dir}")

    if st.button("Run"):
        try:
            fill_form_from_excel(data, username, password)
            st.success("Form filling process completed successfully!")
        except Exception as e:
            st.error(f"An error occurred {e}")