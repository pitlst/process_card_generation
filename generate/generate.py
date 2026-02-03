import datetime
import streamlit as st
import pandas as pd
import json
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from pathlib import Path
from matplotlib import font_manager

font_path = Path(__file__).parent / 'SourceHanSansSC-Normal.otf'
font_manager.fontManager.addfont(font_path)
prop = font_manager.FontProperties(fname=font_path)

plt.rcParams['font.family'] = prop.get_name()
plt.rcParams['axes.unicode_minus'] = False

title = '工序卡生成'
st.set_page_config(page_title=title, layout='wide')
st.title(title)

path = Path(__file__).parent.parent / 'database' / '工序卡模板.json'
pdf_path = Path(__file__).parent.parent / 'source'

if 'res' not in st.session_state:
    st.session_state['res'] = None


def make_main_run(item: dict):
    '''绘图的主逻辑'''
    def check_pdf_file():
        '''检查并删除多余的pdf'''
        reuqest_time = datetime.datetime.now() - datetime.timedelta(hours=1)
        for item_file in pdf_path.iterdir():
            if not item_file.is_file():
                continue
            # 检查是否为 PDF 文件
            if item_file.suffix.lower() != '.pdf':
                continue
            # 获取文件名（不含扩展名）作为时间字符串
            file_time = datetime.datetime.strptime(item_file.stem, '%Y-%m-%d %H:%M:%S')
            if file_time < reuqest_time:
                item_file.unlink()

    # A3大小
    fig, ax = plt.subplots(figsize=(420/25.4, 297/25.4))
    # 隐藏坐标轴，设置绘图范围 0-100 便于定位
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    # 外边框
    ax.add_patch(patches.Rectangle((2, 2), 96, 96, linewidth=1, edgecolor='black', facecolor='none'))
    ax.add_patch(patches.Rectangle((4, 5), 92, 90, linewidth=1, edgecolor='black', facecolor='none'))
    # 左侧密级
    ax.text(6, 95.5, "株机公司普通商密 ▲ 5年", fontsize=14, fontweight='bold', horizontalalignment='left')
    # 右侧工艺代码
    ax.text(94, 95.5, "工艺 22", fontsize=14, fontweight='bold', horizontalalignment='right')
    # 标题
    ax.text(50, 85, "工艺文件", fontsize=48, fontweight='bold', horizontalalignment='center')
    # 产品型号
    ax.text(24, 70, f"产品型号   {item["项目名称"]}", fontsize=24, horizontalalignment='right')
    ax.plot([30, 47], [69, 69], 'k-', linewidth=1.5)
    # 文件名称
    ax.text(61, 70, f"文件名称   {item["工序名称"]}", fontsize=24, horizontalalignment='right')
    ax.plot([64, 80], [69, 69], 'k-', linewidth=1.5)
    # 文件编号
    ax.text(24, 50, f"文件编号   AJP1023290A-22-01", fontsize=24, horizontalalignment='right')
    ax.plot([30, 47], [49, 49], 'k-', linewidth=1.5)

    temp_pdf_name = f'{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.pdf'
    temp_pdf_path = pdf_path / temp_pdf_name
    check_pdf_file()
    plt.savefig(temp_pdf_path, dpi=300, bbox_inches='tight', facecolor='white')
    with open(temp_pdf_path, "rb") as pdf_file:
        pdf_bytes = pdf_file.read()
    return temp_pdf_name, temp_pdf_path, pdf_bytes


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
    if st.session_state['res'] is None:
        with st.container(horizontal=True):
            submit_label = st.button('双击开始生成', icon=':material/send:', shortcut='enter')
            cancel_label = st.button('返回', icon=':material/close:', shortcut='esc')
    else:
        temp_pdf_name, temp_pdf_path, pdf_bytes = st.session_state['res']
        with st.container(horizontal=True):
            submit_label = st.button('双击重新生成', icon=':material/send:', shortcut='enter')
            cancel_label = st.button('返回', icon=':material/close:', shortcut='esc')
            st.download_button(
                label='下载绘制结果',
                data=pdf_bytes,
                file_name=temp_pdf_name,
                mime='application/pdf',
                icon=':material/download:',
            )
        st.pdf(temp_pdf_path, height='stretch')
    if submit_label:
        st.session_state['res'] = make_main_run(temp_config)
    elif cancel_label:
        st.session_state['res'] = None
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
