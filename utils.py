import base64
import streamlit as st


def bg_page(image_file):
    st.markdown(
    """
    <style>
    .stApp {{
        background-image: url("https://raw.githubusercontent.com/jeanaraujoqca/automacao_horas/refs/heads/main/bg_dark.png)";
        background-size: cover
    }}
    </style>
    """,
    unsafe_allow_html=True
    )
