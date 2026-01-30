import streamlit as st

st.sidebar.title("工序卡自动生成服务")

generate_page = st.Page("generate/generate.py", title="工序卡生成", icon=":material/dashboard:", default=True)
maintain_template_page = st.Page("maintain/template.py", title="工序卡模板维护", icon=":material/dashboard:")
maintain_action_page = st.Page("maintain/action.py", title="基础资料_作业动作库维护", icon=":material/dashboard:")
maintain_equipment_page = st.Page("maintain/equipment.py", title="基础资料_工艺装备库维护", icon=":material/dashboard:")
maintain_configuration_page = st.Page("maintain/configuration.py", title="基础资料_构型库维护", icon=":material/dashboard:")
maintain_material_page = st.Page("maintain/material.py", title="基础资料_物料数据库维护", icon=":material/dashboard:")

pg = st.navigation(
    {
        "工序卡生成": [generate_page],
        "基础资料维护": [
            maintain_template_page,
            maintain_action_page,
            maintain_equipment_page,
            maintain_configuration_page,
            maintain_material_page
        ]
    }, 
    position="sidebar"
)
pg.run()
