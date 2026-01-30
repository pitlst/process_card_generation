import streamlit as st
import pandas as pd
from pathlib import Path


title = "工序卡模板维护"
st.set_page_config(page_title=title,layout="wide")
st.title(title)

with st.container(horizontal=True):
    add_label = st.button("新增", icon=':material/add:')
    change_labek = st.button("修改", icon=':material/edit:')
    delete_label = st.button("删除", icon=':material/delete:')


temp_data = pd.DataFrame({
    "模板编码":[],
    "工序编码":[],
    "工序名称":[],
    "适用车型":[],
    "专业分类":[],
})

st.dataframe(temp_data)

@st.dialog("工序卡模板详情")
def detail_view(item: dict):
    ...
    
