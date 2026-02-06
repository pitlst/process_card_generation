import datetime
import streamlit as st
import pandas as pd
import json
from pathlib import Path
from docxtpl import DocxTemplate
from generate.word_api import create_document

title = '工序卡生成'
st.set_page_config(page_title=title, layout='wide')
st.title(title)

path = Path(__file__).parent.parent / 'database' / '工序卡模板.json'
template_path = Path(__file__).parent.parent / 'template' / '工序卡模板.docx'
source_path = Path(__file__).parent.parent / 'source'

def make_main_run(item: dict):
    # 绘图的主逻辑'''
    temp_name = f'{datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.docx'
    temp_path = source_path / temp_name
    create_document(temp_path, item)

    # 检查并删除多余的文档
    reuqest_time = datetime.datetime.now() - datetime.timedelta(minutes=10)
    for item_file in source_path.iterdir():
        if not item_file.is_file():
            continue
        if item_file.suffix.lower() != '.docx':
            continue
        if item_file.name[0: 2] == "~$":
            continue
        file_time = datetime.datetime.strptime(item_file.stem, '%Y-%m-%d-%H-%M-%S')
        if file_time < reuqest_time:
            item_file.unlink()
    # 返回文件的字节流
    with open(temp_path, 'rb') as _file:
        _bytes = _file.read()
    return temp_name, _bytes


@st.cache_data(ttl=3600, show_time=True, scope='session')
def get_template_data() -> dict:
    '''获取本地模板配置文件中的数据'''
    with open(path, mode='r', encoding='utf8') as file:
        return json.loads(file.read())


@st.dialog('生成补充信息', width='large', dismissible=False)
def generate_page(index: int):
    '''生成工序卡需要补充信息的页面'''
    temp_config = get_template_data()[index]
    st.text('这里填写需要你补充的信息')
    with st.container(horizontal=True):
        temp_config['项目名称'] = st.text_input('项目名称')
        temp_config['项目编码'] = st.text_input('项目编码')
        temp_config['密级/保密期限'] = st.selectbox('密级/保密期限', options=['普通商密', '工作秘密'])
    with st.container(horizontal=True):
        temp_config['文件编号'] = st.text_input('文件编号')
        temp_config['零部件图号'] = st.text_input('零部件图号')
    with st.container(horizontal=True):
        temp_config['编制'] = st.text_input('编制')
        temp_config['编制日期'] = st.date_input('编制日期', datetime.datetime.now())
        temp_config['校对'] = st.text_input('校对')
        temp_config['校对日期'] = st.date_input('校对日期', datetime.datetime.now())
    with st.container(horizontal=True):
        temp_config['审核'] = st.text_input('审核')
        temp_config['审核日期'] = st.date_input('审核日期', datetime.datetime.now())
        temp_config['标准化'] = st.text_input('标准化')
        temp_config['标准化日期'] = st.date_input('标准化日期', datetime.datetime.now())
    with st.container(horizontal=True):
        temp_config['会签'] = st.text_input('会签')
        temp_config['会签日期'] = st.date_input('会签日期', datetime.datetime.now())
        temp_config['批准'] = st.text_input('批准')
        temp_config['批准日期'] = st.date_input('批准日期', datetime.datetime.now())
    with st.container(horizontal=True):
        temp_config['失效日期'] = st.date_input('失效日期', datetime.datetime.now() + datetime.timedelta(weeks=48))
        temp_config['文件版本'] = st.text_input('文件版本')

    event = st.data_editor(
        pd.DataFrame(
            {
                '作业顺序': [ch['作业顺序'] for ch in temp_config['工步']],
                '工步名称': [ch['工步名称'] for ch in temp_config['工步']],
                '资质要求': [ch['资质要求'] for ch in temp_config['工步']],
                '注意内容': [ch['注意内容'] for ch in temp_config['工步']],
                '是否关键工步': [ch['是否关键工步'] for ch in temp_config['工步']],
                '是否特殊过程': [ch['是否特殊过程'] for ch in temp_config['工步']],
                '是否八防工序': [ch['是否八防工序'] for ch in temp_config['工步']],
                '是否五防工序': [ch['是否五防工序'] for ch in temp_config['工步']],
                '是否关键质量控制点': [ch['是否关键质量控制点'] for ch in temp_config['工步']],
                '工艺装备': [ch['工艺装备'] for ch in temp_config['工步']],
            }
        ),
        hide_index=True
    )
    st.info('对应的生成记录会在后台保存10分钟，找回请检查后台文件中的source文件夹')
    temp_empty = st.empty()
    with temp_empty:
        with st.container(horizontal=True):
            submit_label = st.button('开始生成', icon=':material/send:', shortcut='enter')
            cancel_label = st.button('返回', icon=':material/close:', shortcut='esc', key='cancel_0')
    if submit_label:
        with st.spinner("正在生成文档中", show_time=True):
            temp_name, docx_bytes = make_main_run(temp_config)
        temp_empty.empty()
        with st.container(horizontal=True):
            cancel_label = st.button('返回', icon=':material/close:', shortcut='esc', key='cancel_1')
            st.download_button(
                label='下载绘制结果',
                data=docx_bytes,
                file_name=temp_name,
                mime='application/docx',
                icon=':material/download:',
            )
    elif cancel_label:
        st.rerun()


st.markdown('##### 选择你要生成工序卡的对应模板')
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
    else:
        generate_page(event.selection.rows[0])  # type: ignore
