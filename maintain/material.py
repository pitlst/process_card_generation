import streamlit as st
from pathlib import Path
from maintain import page_make

title = '物料基础资料维护'
st.set_page_config(page_title=title, layout='wide')
st.title(title)

path = Path(__file__).parent.parent / 'database' / '物料基础资料.csv'

page_make(path)
