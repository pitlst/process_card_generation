import matplotlib.pyplot as plt
from pathlib import Path
from matplotlib import font_manager

font_path = Path(__file__).parent / 'SourceHanSansSC-Normal.otf'
font_manager.fontManager.addfont(font_path)
prop = font_manager.FontProperties(fname=font_path)

plt.rcParams['font.family'] = prop.get_name()
plt.rcParams['axes.unicode_minus'] = False


def draw_process_card():
    width = 420
    height = 297
    fig, ax = plt.subplots(figsize=(16.54, 11.69))
    ax.set_xlim(0, width)
    ax.set_ylim(0, height)
    ax.axis('off')

    # 配色
    c_border = '#000000'
    c_header = '#D3D3D3'
    c_label = '#F0F0F0'
    ...
