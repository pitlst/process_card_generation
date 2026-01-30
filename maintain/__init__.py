import streamlit as st
import pandas as pd
from pathlib import Path

def page_make(path: Path):
    '''通用的维护页面生成'''
    
    @st.cache_data(ttl=3600)
    def get_data():
        '''获取表格中的数据'''
        return pd.read_csv(path, encoding='utf-8')


    @st.cache_data
    def convert_for_download(df):
        return df.to_csv().encode('utf-8')


    num_rows = 'fixed'
    with st.container(horizontal=True):
        with st.container(width="content"):
            temp_data = get_data()
            st.download_button(
                label='下载批量新增使用的excel模板',
                data=convert_for_download(temp_data.drop(temp_data.index)),
                file_name='模板.csv',
                mime='text/csv',
                icon=':material/download:',
            )
            refresh_label = st.button('手动刷新页面缓存', icon=':material/refresh:')
        uploaded_file = st.file_uploader('**上传Excel替换或新增**', type=['csv, xlsx, xls'])
    with st.container(horizontal=True):
        save_label = st.button('保存到服务器后台中', icon=':material/save:')
        if st.toggle('启用新增(会导致排序功能失效，不影响修改)'):
            num_rows = 'dynamic'

    change_df = st.data_editor(get_data(), height='content', num_rows=num_rows)
    
    if save_label:
        change_df.to_csv(path, encoding='utf-8', index=False)
        get_data.clear()
        st.toast('保存成功', icon='🎉')
    if refresh_label:
        get_data.clear()
        st.toast('缓存刷新成功', icon='🎉')
    if not uploaded_file is None:
        
        ...
