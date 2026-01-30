import streamlit as st
from pathlib import Path
from maintain import page_make

title = "构型与设计方案项基础资料维护"
st.set_page_config(page_title=title,layout="wide")
st.title(title)

path = Path(__file__).parent.parent / 'database' / '构型与设计方案项基础资料.csv'

page_make(path)