import streamlit as st
import pandas as pd
import numpy as np




text_contents = '''This is some text'''
st.download_button('Download some text', text_contents)

binary_contents = b'example content'
# Defaults to 'application/octet-stream'
st.download_button('Download binary file', binary_contents)

with open("pages/amplopgabung.docx", "rb") as file:
    btn = st.download_button(
        label="Download image",
        data=file,
        file_name="amplopgabung.docx",
        mime="image/png"
    )