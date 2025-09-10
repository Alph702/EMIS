To read an Excel sheet uploaded with st.file_uploader and access its filename, you can use the following approach:

```python
import streamlit as st
import pandas as pd

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Read the Excel file into a DataFrame
    df = pd.read_excel(uploaded_file)
    st.write(df)
    
    # Get the filename
    st.write("Filename:", uploaded_file.name)
```

- The uploaded_file object returned by st.file_uploader has a .name attribute that gives you the filename.
- You can pass uploaded_file directly to pd.read_excel, as it is a file-like object.

This method allows you to both read the Excel file and display or use its filename in your Streamlit app. For more details, see the official documentation: [How do you retrieve the filename of a file uploaded with st.file_uploader?](https://docs.streamlit.io/knowledge-base/using-streamlit/retrieve-filename-uploaded) and [st.file_uploader API reference](https://docs.streamlit.io/develop/api-reference/widgets/st.file_uploader).