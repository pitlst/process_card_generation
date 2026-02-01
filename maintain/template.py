import streamlit as st
import pandas as pd
import json
from typing import Any
from pathlib import Path

root_path = Path(__file__).parent.parent / 'database'
path = root_path / '工序卡模板.json'
action_path = root_path / '作业动作库基础资料.csv'
configuration_poath = root_path / '构型与设计方案项基础资料.csv'
equipment_path = root_path / '工艺装备基础资料.csv'
material_path = root_path / '物料基础资料.csv'


@st.cache_data(ttl=3600, show_time=True, scope="session")
def get_template_data() -> dict[str, Any]:
    '''获取本地模板配置文件中的数据'''
    with open(path, mode="r", encoding="utf8") as file:
        return json.loads(file.read())


@st.cache_data(ttl=3600, show_time=True, scope="session")
def get_total_configuration():
    '''获取所有目前的设计方案项'''
    configuration_data = pd.read_csv(configuration_poath, encoding='utf-8')
    return list(configuration_data["设计方案项编码"])


@st.cache_data(ttl=3600, show_time=True, scope="session")
def get_total_equipment():
    '''获取所有目前的工艺装备'''
    configuration_data = pd.read_csv(equipment_path, encoding='utf-8')
    return list(configuration_data["工艺装备编码"])


@st.cache_data(ttl=3600, show_time=True, scope="session")
def get_total_material():
    '''获取所有目前的物料'''
    configuration_data = pd.read_csv(material_path, encoding='utf-8')
    return list(configuration_data["物料编码"])


@st.cache_data(ttl=3600, show_time=True, scope="session")
def get_total_action():
    '''获取所有目前的工作'''
    configuration_data = pd.read_csv(action_path, encoding='utf-8')
    return list(configuration_data["作业动作编码"])


def get_template() -> dict[str, Any]:
    '''获取一个没有数据的纯模板配置文件'''
    # 单据头
    template_config = {}
    template_config["模板编码"] = None
    template_config["工序编码"] = None
    template_config["工序名称"] = None
    template_config["适用车型"] = None
    template_config["专业分类"] = None
    template_config["设计方案项"] = None
    template_config["工步"] = []
    template_config["物料清单"] = []
    # 工步分录
    template_workstep = {}
    template_workstep["作业顺序"] = None
    template_workstep["工步名称"] = None
    template_workstep["资质要求"] = None
    template_workstep["注意内容"] = None
    template_workstep["附件图片"] = None
    template_workstep["是否关键工步"] = None
    template_workstep["是否特殊过程"] = None
    template_workstep["是否八防工序"] = None
    template_workstep["是否五防工序"] = None
    template_workstep["是否关键质量控制点"] = None
    template_workstep["对应动作"] = []
    template_workstep["对应工艺装备"] = []
    # 工步对应动作
    template_workstep_action = {}
    template_workstep_action["作业动作编码"] = None
    template_workstep_action["工艺参数要求"] = None
    template_workstep_action["验证形式"] = None
    template_workstep_action["验证结果"] = None
    template_workstep["对应动作"].append(template_workstep_action)
    # 工步对应工艺装备
    template_workstep_equipment = {}
    template_workstep_equipment["工艺装备编码"] = None
    template_workstep["对应工艺装备"].append(template_workstep_equipment)
    template_config["工步"].append(template_workstep)
    # 物料分录
    template_material = {}
    template_material["物料编码"] = None
    template_material["物料数量"] = None
    template_material["是否关键物料"] = None
    template_material["是否不装车辅料"] = None
    template_config["物料清单"].append(template_material)
    return template_config


title = "工序卡模板维护"
st.set_page_config(page_title=title, layout="wide")
st.title(title)

with st.container(horizontal=True):
    add_label = st.button("新增", icon=':material/add:')
    change_labek = st.button("修改", icon=':material/edit:')
    delete_label = st.button("删除", icon=':material/delete:')
    refresh_label = st.button('刷新', icon=':material/refresh:')

local_data = get_template_data()
temp_data = pd.DataFrame({
    "模板编码": [item["模板编码"] for item in local_data],  # pyright: ignore[reportArgumentType]
    "工序编码": [item["工序编码"] for item in local_data],  # pyright: ignore[reportArgumentType]
    "工序名称": [item["工序名称"] for item in local_data],  # pyright: ignore[reportArgumentType]
    "适用车型": [item["适用车型"] for item in local_data],  # pyright: ignore[reportArgumentType]
    "专业分类": [item["专业分类"] for item in local_data],  # pyright: ignore[reportArgumentType]
})

st.dataframe(temp_data, hide_index=True, on_select="rerun", selection_mode="single-row")

temp_template_config = {}
temp_template_config["工序编码"] = None


@st.dialog("工序卡模板详情", width="large", dismissible=False)
def detail_view(item: dict):
    st.text("模板单据头信息")
    with st.container(horizontal=True):
        item["模板编码"] = st.text_input("模板编码", value=item["模板编码"])
        item["工序编码"] = st.text_input("工序编码", value=item["工序编码"])
        item["工序名称"] = st.text_input("工序名称", value=item["工序名称"])
        item["适用车型"] = st.text_input("适用车型", value=item["适用车型"])
        item["专业分类"] = st.text_input("专业分类", value=item["专业分类"])
        total_configuration = get_total_configuration()
        if item["设计方案项"] in total_configuration:
            item["设计方案项"] = st.selectbox("设计方案项", get_total_configuration(), total_configuration.index(item["设计方案项"]))
        else:
            item["设计方案项"] = st.selectbox("设计方案项", get_total_configuration())
    st.text("模板的工步分录")
    if 1:
        temp_df = pd.DataFrame(
            {
                "作业顺序": [ch["作业顺序"] for ch in item["工步"]],
                "工步名称": [ch["工步名称"] for ch in item["工步"]],
                "资质要求": [ch["资质要求"] for ch in item["工步"]],
                "注意内容": [ch["注意内容"] for ch in item["工步"]],
                "是否关键工步": [ch["是否关键工步"] for ch in item["工步"]],
                "是否特殊过程": [ch["是否特殊过程"] for ch in item["工步"]],
                "是否八防工序": [ch["是否八防工序"] for ch in item["工步"]],
                "是否五防工序": [ch["是否五防工序"] for ch in item["工步"]],
                "是否关键质量控制点": [ch["是否关键质量控制点"] for ch in item["工步"]],
                "对应工艺装备": [[_ch["工艺装备编码"] for _ch in ch["对应工艺装备"]] for ch in item["工步"]],
            }
        )
        edited_df = st.data_editor(
            temp_df,
            column_config={
                "作业顺序": st.column_config.NumberColumn(
                    "作业顺序",
                    help="表示工步的执行顺序，由小到大执行，最先执行的工步为0",
                    min_value=0,
                    max_value=1000,
                    step=1,
                    default=0
                ),
                "是否关键工步": st.column_config.CheckboxColumn(
                    "是否关键工步",
                    default=False),
                "是否特殊过程": st.column_config.CheckboxColumn(
                    "是否特殊过程",
                    default=False),
                "是否八防工序": st.column_config.CheckboxColumn(
                    "是否八防工序",
                    default=False),
                "是否五防工序": st.column_config.CheckboxColumn(
                    "是否五防工序",
                    default=False),
                "是否关键质量控制点": st.column_config.CheckboxColumn(
                    "是否关键质量控制点",
                    default=False),
            },
            hide_index=True)
    with st.container(horizontal=True):
        if st.button("提交"):
            st.rerun()
        if st.button("取消"):
            st.rerun()


if refresh_label:
    get_template_data.clear()
    get_total_configuration.clear()
    get_total_equipment.clear()
    get_total_material.clear()
    get_total_action.clear()
    st.toast('刷新成功', icon='🎉')

if add_label:
    detail_view(get_template())
elif change_labek:
    ...
elif delete_label:
    ...
