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


@st.cache_data(ttl=3600, show_time=True, scope='session')
def get_template_data() -> dict:
    '''获取本地模板配置文件中的数据'''
    with open(path, mode='r', encoding='utf8') as file:
        return json.loads(file.read())


@st.cache_data(ttl=3600, show_time=True, scope='session')
def get_total_configuration() -> list[str]:
    '''获取所有目前的设计方案项'''
    configuration_data = pd.read_csv(configuration_poath, encoding='utf-8')
    return list(configuration_data['设计方案项编码'])


@st.cache_data(ttl=3600, show_time=True, scope='session')
def get_total_equipment() -> list[str]:
    '''获取所有目前的工艺装备'''
    configuration_data = pd.read_csv(equipment_path, encoding='utf-8')
    return list(configuration_data['工艺装备编码'])


@st.cache_data(ttl=3600, show_time=True, scope='session')
def get_total_material() -> list[str]:
    '''获取所有目前的物料'''
    configuration_data = pd.read_csv(material_path, encoding='utf-8')
    return list(configuration_data['物料编码'])


@st.cache_data(ttl=3600, show_time=True, scope='session')
def get_total_action() -> list[str]:
    '''获取所有目前的工作'''
    configuration_data = pd.read_csv(action_path, encoding='utf-8')
    return list(configuration_data['作业动作编码'])


def get_template() -> dict[str, Any]:
    '''获取一个没有数据的纯模板配置文件'''
    template_config = {}
    template_config['模板编码'] = None
    template_config['工序编码'] = None
    template_config['工序名称'] = None
    template_config['适用车型'] = None
    template_config['专业分类'] = None
    template_config['设计方案项'] = None
    template_config['工步'] = []
    template_config['物料清单'] = []
    return template_config


def get_workstep_template() -> dict[str, Any]:
    '''获取一个没有数据的纯模板工步配置文件'''
    template_workstep = {}
    template_workstep['作业顺序'] = None
    template_workstep['工步名称'] = None
    template_workstep['资质要求'] = None
    template_workstep['注意内容'] = None
    template_workstep['附件图片'] = None
    template_workstep['是否关键工步'] = None
    template_workstep['是否特殊过程'] = None
    template_workstep['是否八防工序'] = None
    template_workstep['是否五防工序'] = None
    template_workstep['是否关键质量控制点'] = None
    template_workstep['动作'] = []
    template_workstep['工艺装备'] = []
    return template_workstep


def get_workstep_action_template() -> dict[str, Any]:
    '''获取一个没有数据的纯模板工步对应动作配置文件'''
    template_workstep_action = {}
    template_workstep_action['作业动作编码'] = None
    template_workstep_action['工艺参数要求'] = None
    template_workstep_action['验证形式'] = None
    template_workstep_action['验证结果'] = None
    return template_workstep_action


# ------------------------------------------
#  其他弹出页面定义的开发
#  MARK: 弹出页面定义
# ------------------------------------------


@st.dialog('工序卡模板新增/修改详情', width='large', dismissible=False)
def detail_view():
    copy_data = st.session_state['page_item']
    st.text('模板单据头信息')
    with st.container(horizontal=True):
        st.session_state['page_item']['模板编码'] = st.text_input('模板编码', value=st.session_state['page_item']['模板编码'])
        st.session_state['page_item']['工序编码'] = st.text_input('工序编码', value=st.session_state['page_item']['工序编码'])
        st.session_state['page_item']['工序名称'] = st.text_input('工序名称', value=st.session_state['page_item']['工序名称'])
    with st.container(horizontal=True):
        st.session_state['page_item']['适用车型'] = st.text_input('适用车型', value=st.session_state['page_item']['适用车型'])
        st.session_state['page_item']['专业分类'] = st.text_input('专业分类', value=st.session_state['page_item']['专业分类'])
        total_configuration = get_total_configuration()
        index = 0
        if st.session_state['page_item']['设计方案项'] in total_configuration:
            index = total_configuration.index(st.session_state['page_item']['设计方案项'])
        st.session_state['page_item']['设计方案项'] = st.selectbox('设计方案项', total_configuration, index)
    total_material = get_total_material()
    default = None
    if set(st.session_state['page_item']['物料清单']) <= set(total_material):
        default = st.session_state['page_item']['物料清单']
    st.session_state['page_item']['物料清单'] = st.multiselect('物料清单', total_material, default)

    st.text('模板的工步分录')
    temp_df = pd.DataFrame(
        {
            '作业顺序': [ch['作业顺序'] for ch in st.session_state['page_item']['工步']],
            '工步名称': [ch['工步名称'] for ch in st.session_state['page_item']['工步']],
            '资质要求': [ch['资质要求'] for ch in st.session_state['page_item']['工步']],
            '注意内容': [ch['注意内容'] for ch in st.session_state['page_item']['工步']],
            '是否关键工步': [ch['是否关键工步'] for ch in st.session_state['page_item']['工步']],
            '是否特殊过程': [ch['是否特殊过程'] for ch in st.session_state['page_item']['工步']],
            '是否八防工序': [ch['是否八防工序'] for ch in st.session_state['page_item']['工步']],
            '是否五防工序': [ch['是否五防工序'] for ch in st.session_state['page_item']['工步']],
            '是否关键质量控制点': [ch['是否关键质量控制点'] for ch in st.session_state['page_item']['工步']],
            '工艺装备': [ch['工艺装备'] for ch in st.session_state['page_item']['工步']],
        }
    )
    with st.container(horizontal=True):
        entry_add_label = st.button('新增', icon=':material/add:', shortcut='alt++shift+w')
        entry_change_label = st.button('修改', icon=':material/edit:', shortcut='alt++shift+e')
        entry_delete_label = st.button('删除', icon=':material/delete:', shortcut='alt++shift+d')
    event = st.dataframe(
        temp_df,
        column_config={
            '作业顺序': st.column_config.NumberColumn(
                '作业顺序',
                help='表示工步的执行顺序，由小到大执行，最先执行的工步为0',
                min_value=0,
                max_value=1000,
                step=1,
                default=0
            ),
            '是否关键工步': st.column_config.CheckboxColumn(
                '是否关键工步',
                default=False),
            '是否特殊过程': st.column_config.CheckboxColumn(
                '是否特殊过程',
                default=False),
            '是否八防工序': st.column_config.CheckboxColumn(
                '是否八防工序',
                default=False),
            '是否五防工序': st.column_config.CheckboxColumn(
                '是否五防工序',
                default=False),
            '是否关键质量控制点': st.column_config.CheckboxColumn(
                '是否关键质量控制点',
                default=False),
        },
        on_select='rerun',
        selection_mode='single-row',
        hide_index=True)
    with st.container(horizontal=True):
        submit_label = st.button('提交', icon=':material/send:', shortcut='enter')
        cancel_label = st.button('取消', icon=':material/close:', shortcut='esc')
    # ------------------------------------------
    #  工序卡页面标志按钮处理
    #  MARK: 标志位按钮处理
    # ------------------------------------------
    if entry_add_label:
        st.session_state['page_workstep_item'] = get_workstep_template()
        st.session_state['page_path'] = 'workstep'
        st.rerun()
    elif entry_change_label:
        if len(event.selection.rows) == 0:  # type: ignore
            st.toast(f'未选择任何行无法修改', icon='🚨')
        else:
            st.session_state['page_workstep_item'] = st.session_state['page_item']['工步'][event.selection.rows[0]]  # type: ignore
            del st.session_state['page_item']['工步'][event.selection.rows[0]]  # type: ignore
            st.session_state['page_path'] = 'workstep'
            st.rerun()
    elif entry_delete_label:
        if len(event.selection.rows) == 0:  # type: ignore
            st.toast(f'未选择任何行无法修改', icon='🚨')
        else:
            del st.session_state['page_item']['工步'][event.selection.rows[0]]  # type: ignore
            st.toast(f'删除成功', icon='🎉')
            st.rerun()
    elif submit_label:
        local_data = list(get_template_data())
        local_data.append(st.session_state['page_item'])
        with open(path, mode='w', encoding='utf8') as file:
            file.write(json.dumps(local_data, indent=4, ensure_ascii=False, default=str))
        get_template_data.clear()
        st.session_state['page_path'] = ''
        st.rerun()
    elif cancel_label:
        local_data = list(get_template_data())
        local_data.append(copy_data)
        with open(path, mode='w', encoding='utf8') as file:
            file.write(json.dumps(local_data, indent=4, ensure_ascii=False, default=str))
        get_template_data.clear()
        st.session_state['page_path'] = ''
        st.rerun()


@st.dialog('工序卡模板新增/修改详情---工步', width='large', dismissible=False)
def detail_workstep_view():
    copy_data = st.session_state['page_workstep_item']
    with st.container(horizontal=True):
        st.session_state['page_workstep_item']['作业顺序'] = st.number_input('作业顺序', min_value=0, max_value=1000, step=1, value=st.session_state['page_workstep_item']['作业顺序'])
        st.session_state['page_workstep_item']['工步名称'] = st.text_input('工步名称', st.session_state['page_workstep_item']['工步名称'])
        st.session_state['page_workstep_item']['资质要求'] = st.text_input('资质要求', st.session_state['page_workstep_item']['资质要求'])
    st.session_state['page_workstep_item']['注意内容'] = st.text_area('注意内容', st.session_state['page_workstep_item']['注意内容'])
    with st.container(horizontal=True):
        st.session_state['page_workstep_item']['是否关键工步'] = st.checkbox("是否关键工步", value=st.session_state['page_workstep_item']['是否关键工步'])
        st.session_state['page_workstep_item']['是否特殊过程'] = st.checkbox("是否特殊过程", value=st.session_state['page_workstep_item']['是否特殊过程'])
        st.session_state['page_workstep_item']['是否八防工序'] = st.checkbox("是否八防工序", value=st.session_state['page_workstep_item']['是否八防工序'])
        st.session_state['page_workstep_item']['是否五防工序'] = st.checkbox("是否五防工序", value=st.session_state['page_workstep_item']['是否五防工序'])
        st.session_state['page_workstep_item']['是否关键质量控制点'] = st.checkbox("是否关键质量控制点", value=st.session_state['page_workstep_item']['是否关键质量控制点'])

    total_equipment = get_total_equipment()
    default = None
    if set(st.session_state['page_workstep_item']['工艺装备']) <= set(total_equipment):
        default = st.session_state['page_workstep_item']['工艺装备']
    st.session_state['page_workstep_item']['工艺装备'] = st.multiselect('工艺装备', total_equipment, default)

    st.text('模板工步的动作分录')
    temp_df = pd.DataFrame(
        {
            '作业动作编码': [ch['作业动作编码'] for ch in st.session_state['page_workstep_item']['动作']],
            '工艺参数要求': [ch['工艺参数要求'] for ch in st.session_state['page_workstep_item']['动作']],
            '验证形式': [ch['验证形式'] for ch in st.session_state['page_workstep_item']['动作']],
            '验证结果': [ch['验证结果'] for ch in st.session_state['page_workstep_item']['动作']],
        }
    )
    with st.container(horizontal=True):
        entry_add_label = st.button('新增', icon=':material/add:', shortcut='alt++shift+w')
        entry_change_label = st.button('修改', icon=':material/edit:', shortcut='alt++shift+e')
        entry_delete_label = st.button('删除', icon=':material/delete:', shortcut='alt++shift+d')
    event = st.dataframe(temp_df, on_select='rerun', selection_mode='single-row', hide_index=True)

    with st.container(horizontal=True):
        submit_label = st.button('提交', icon=':material/send:', shortcut='enter')
        cancel_label = st.button('取消', icon=':material/close:', shortcut='esc')

    # ------------------------------------------
    #  工步页面标志按钮处理
    #  MARK: 工步页面标志按钮处理
    # ------------------------------------------
    if entry_add_label:
        st.session_state['page_workstep_action_item'] = get_workstep_action_template()
        st.session_state['page_path'] = 'workstep_action'
        st.rerun()
    elif entry_change_label:
        if len(event.selection.rows) == 0:  # type: ignore
            st.toast(f'未选择任何行无法修改', icon='🚨')
        else:
            st.session_state['page_workstep_action_item'] = st.session_state['page_workstep_item']['动作'][event.selection.rows[0]]  # type: ignore
            del st.session_state['page_workstep_item']['动作'][event.selection.rows[0]]  # type: ignore
            st.session_state['page_path'] = 'workstep_action'
            st.rerun()
    elif entry_delete_label:
        if len(event.selection.rows) == 0:  # type: ignore
            st.toast(f'未选择任何行无法修改', icon='🚨')
        else:
            del st.session_state['page_workstep_item']['动作'][event.selection.rows[0]]  # type: ignore
            st.toast(f'删除成功', icon='🎉')
            st.rerun()
    elif submit_label:
        st.session_state['page_item']['工步'].append(st.session_state['page_workstep_item'])
        st.session_state['page_path'] = 'main'
        st.rerun()
    elif cancel_label:
        st.session_state['page_item']['工步'].append(copy_data)
        st.session_state['page_path'] = 'main'
        st.rerun()


@st.dialog('工序卡模板新增/修改详情---工步对应动作', width='large', dismissible=False)
def detail_workstep_action_view():
    copy_data = st.session_state['page_workstep_action_item']
    with st.container(horizontal=True):
        total_data = get_total_action()
        index = 0
        if st.session_state['page_workstep_action_item']['作业动作编码'] in total_data:
            index = total_data.index(st.session_state['page_workstep_action_item']['作业动作编码'])
        st.session_state['page_workstep_action_item']['作业动作编码'] = st.selectbox('作业动作编码', total_data, index)
        st.session_state['page_workstep_action_item']['工艺参数要求'] = st.text_input('工艺参数要求', value=st.session_state['page_workstep_action_item']['工艺参数要求'])
    with st.container(horizontal=True):
        total_data = ['定量', '定性']
        index = 0
        if st.session_state['page_workstep_action_item']['验证形式'] in total_data:
            index = total_data.index(st.session_state['page_workstep_action_item']['验证形式'])
        st.session_state['page_workstep_action_item']['验证形式'] = st.selectbox('验证形式', total_data, index)
        if st.session_state['page_workstep_action_item']['验证形式'] == '定量':
            total_data_2 = ['合格', '不合格']
            index = 0
            if st.session_state['page_workstep_action_item']['验证结果'] in total_data_2:
                index = total_data_2.index(st.session_state['page_workstep_action_item']['验证结果'])
            st.session_state['page_workstep_action_item']['验证结果'] = st.selectbox('验证结果', total_data_2, index)
        else:
            st.session_state['page_workstep_action_item']['验证结果'] = st.text_input('验证结果', value=st.session_state['page_workstep_action_item']['验证结果'])
    with st.container(horizontal=True):
        submit_label = st.button('提交', icon=':material/send:', shortcut='enter')
        cancel_label = st.button('取消', icon=':material/close:', shortcut='esc')
    if submit_label:
        st.session_state['page_workstep_item']['动作'].append(st.session_state['page_workstep_action_item'])
        st.session_state['page_path'] = 'workstep'
        st.rerun()
    if cancel_label:
        st.session_state['page_workstep_item']['动作'].append(copy_data)
        st.session_state['page_path'] = 'workstep'
        st.rerun()


# ------------------------------------------
#  主页面定义的开始
#  MARK: 主页面定义
# ------------------------------------------


def main():
    title = '工序卡模板维护'
    st.set_page_config(page_title=title, layout='wide')
    st.title(title)

    st.info("少量的维护可以直接在页面更改，大量更新建议下载模板进行更新，模板中会带有现有的数据，因为开发周期，目前没有做excel的处理，需要将excel导出为csv才能上传")
    st.warning("目前没有做多人同时操作的隔离，所以需要注意维护数据时的冲突问题")

    with st.container(horizontal=True):
        add_label = st.button('新增', icon=':material/add:', shortcut='alt+w')
        change_label = st.button('修改', icon=':material/edit:', shortcut='alt+e')
        delete_label = st.button('删除', icon=':material/delete:', shortcut='alt+d')
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

    # ------------------------------------------
    #  标志位按钮处理
    #  MARK: 标志位按钮处理
    # ------------------------------------------
    if refresh_label:
        get_template_data.clear()
        get_total_configuration.clear()
        get_total_equipment.clear()
        get_total_material.clear()
        get_total_action.clear()
        st.rerun()

    if add_label:
        st.session_state['page_item'] = get_template()
        st.session_state['page_path'] = 'main'
        detail_view()
    elif change_label:
        if len(event.selection.rows) == 0:  # type: ignore
            st.toast(f'未选择任何行无法修改', icon='🚨')
        else:
            st.session_state['page_item'] = local_data[event.selection.rows[0]]  # type: ignore
            del local_data[event.selection.rows[0]]  # type: ignore
            with open(path, mode='w', encoding='utf8') as file:
                file.write(json.dumps(local_data, indent=4, ensure_ascii=False, default=str))
            st.session_state['page_path'] = 'main'
            get_template_data.clear()
            detail_view()
    elif delete_label:
        if len(event.selection.rows) == 0:  # type: ignore
            st.toast(f'未选择任何行无法修改', icon='🚨')
        else:
            del local_data[event.selection.rows[0]]  # type: ignore
            with open(path, mode='w', encoding='utf8') as file:
                file.write(json.dumps(local_data, indent=4, ensure_ascii=False, default=str))
            get_template_data.clear()
            st.rerun()


# dialog的路由页面参数存储初始化
if 'page_path' not in st.session_state:
    st.session_state['page_path'] = ''
# 页面修改的对应单据的id初始化
if 'page_item' not in st.session_state:
    st.session_state['page_item'] = get_template()
if 'page_workstep_item' not in st.session_state:
    st.session_state['page_workstep_item'] = get_workstep_template()
if 'page_workstep_action_item' not in st.session_state:
    st.session_state['page_workstep_action_item'] = get_workstep_action_template()

# ------------------------------------------
#  弹出页面路由处理
#  MARK: 弹出页面路由处理
# ------------------------------------------
if st.session_state['page_path'] == '':
    main()
elif st.session_state['page_path'] == 'main':
    main()
    detail_view()
elif st.session_state['page_path'] == 'workstep':
    main()
    detail_workstep_view()
elif st.session_state['page_path'] == 'workstep_action':
    main()
    detail_workstep_action_view()
