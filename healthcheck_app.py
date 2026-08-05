import sys
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Healthcheck", layout="wide")
st.title("Healthcheck Streamlit Cloud")
st.success("Ambiente carregado corretamente.")
st.code(f"Python: {sys.version}")
st.code(f"Streamlit: {st.__version__}")
st.code(f"Pandas: {pd.__version__}")
