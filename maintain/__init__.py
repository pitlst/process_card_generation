import streamlit as st
import pandas as pd
from pathlib import Path


@st.cache_data(ttl=3600, show_time=True, scope="session")
def get_data(path: Path):
    '''获取本地csv中的数据'''
    return pd.read_csv(path, encoding='utf-8')


def page_make(path: Path):
    '''通用的基础资料维护页面生成'''
    st.info("少量的维护可以直接在页面更改，大量更新建议下载模板进行更新，模板中会带有现有的数据，因为开发周期，目前没有做excel的处理，需要将excel导出为csv才能上传")
    st.warning("目前没有做多人同时操作的隔离，所以需要注意维护数据时的冲突问题")
    num_rows = 'fixed'
    temp_data = get_data(path)
    with st.container(horizontal=True):
        with st.container(width="content"):
            st.download_button(
                label='下载批量更新模板',
                data=temp_data.to_csv().encode('utf-8'),
                file_name='模板.csv',
                mime='text/csv',
                icon=':material/download:',
            )
            refresh_label = st.button('手动刷新页面缓存', icon=':material/refresh:')
            save_label = st.button('保存到后台中', icon=':material/save:')

        with st.container():
            uploaded_file = st.file_uploader('**上传批量更新的数据**', type=['csv'])
            add_label = st.toggle('启用新增(会导致排序功能失效，不影响修改)')

    if add_label:
        num_rows = 'dynamic'
    change_df = st.data_editor(get_data(path), height='content', num_rows=num_rows, hide_index=True)

    if save_label:
        change_df.to_csv(path, encoding='utf-8', index=False)
        get_data.clear()
        st.toast('保存成功', icon='🎉')
    if refresh_label:
        get_data.clear()
        st.toast('缓存刷新成功', icon='🎉')
    if not uploaded_file is None:
        pd.read_csv(uploaded_file, encoding='utf-8').to_csv(path, encoding='utf-8', index=False)
        get_data.clear()
        st.toast('更新数据成功', icon='🎉')
