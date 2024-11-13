import base64
import streamlit as st


def bg_page(image_file):
    with open(image_file, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read())
    st.markdown(
    f"""
    <style>
    .stApp {{
        background-image: url("{https://raw.githubusercontent.com/jeanaraujoqca/automacao_horas/refs/heads/main/bg_dark.png})";
        background-size: cover
    }}
    </style>
    """,
    unsafe_allow_html=True
    )
