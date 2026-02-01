import streamlit as st
import pandas as pd
import json
from pathlib import Path

title = "工序卡生成"
st.set_page_config(page_title=title, layout="wide")
st.title(title)

path = Path(__file__).parent.parent / 'database' / '工序卡模板.json'


@st.cache_data(ttl=3600, show_time=True, scope='session')
def get_template_data() -> dict:
    '''获取本地模板配置文件中的数据'''
    with open(path, mode='r', encoding='utf8') as file:
        return json.loads(file.read())

@st.dialog('生成补充信息', width='large', dismissible=False)
def generate_page(index: int):
    '''生成工序卡需要补充信息的页面'''
    temp_config = get_template_data()[index]
    temp_config[] = st.text_input("")
    st.text("生成完成")


st.markdown("##### 选择你要生成工序卡的对应模板")
with st.container(horizontal=True):
    generate_label = st.button('生成', icon=':material/build:', shortcut='alt+g')
    refresh_label = st.button('刷新', icon=':material/refresh:', shortcut='alt+f')
local_data = get_template_data()
temp_data = pd.DataFrame({
    '模板编码': [item['模板编码'] for item in local_data],  # pyright: ignore[reportArgumentType]
    '工序编码': [item['工序编码'] for item in local_data],  # pyright: ignore[reportArgumentType]
    '工序名称': [item['工序名称'] for item in local_data],  # pyright: ignore[reportArgumentType]
    '适用车型': [item['适用车型'] for item in local_data],  # pyright: ignore[reportArgumentType]
    '专业分类': [item['专业分类'] for item in local_data],  # pyright: ignore[reportArgumentType]
})
event = st.dataframe(temp_data, hide_index=True, on_select='rerun', selection_mode='single-row')

if refresh_label:
    get_template_data.clear()
elif generate_label:
    if len(event.selection.rows) == 0:  # type: ignore
        st.toast(f'未选择任何行无法修改', icon='🚨')
    generate_page(event.selection.rows[0]) # type: ignore