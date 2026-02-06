import win32com.client
from pathlib import Path

CM_TO_POINT = 28.35

def create_document(file_path: Path, item: dict):
    # 获取绝对路径，Word COM 接口通常需要绝对路径
    file_path = file_path.absolute()
    # 检查文件是否已存在，如果存在先删除，避免 SaveAs 弹窗或报错
    if file_path.exists():
        file_path.unlink()
    word = win32com.client.Dispatch("Word.Application") # 启动 Word 应用程序
    word.Visible = True  # 设置为 True 可以看到 Word 界面，方便调试；设置为 False 则后台运行
    word.DisplayAlerts = 0  # 禁用警告弹窗
    doc = word.Documents.Add() # 新建文档
    
    # 设置页面布局：A4 横向
    # wdOrientLandscape = 1, wdPaperA4 = 7
    doc.PageSetup.Orientation = 1
    doc.PageSetup.PaperSize = 7

    # 设置页边距：左 2.2cm，上下右 0.5cm
    # 直接计算点数 (1 cm ≈ 28.35 points)，避免调用 word.CentimetersToPoints 可能出现的 COM 错误
    
    doc.PageSetup.LeftMargin = 2.2 * CM_TO_POINT
    doc.PageSetup.RightMargin = 0.5 * CM_TO_POINT
    doc.PageSetup.TopMargin = 0.5 * CM_TO_POINT
    doc.PageSetup.BottomMargin = 0.5 * CM_TO_POINT

    # 绘制页面边框矩形（沿着页边距）
    # 获取页面宽高和边距
    page_width = doc.PageSetup.PageWidth
    page_height = doc.PageSetup.PageHeight
    left = doc.PageSetup.LeftMargin
    top = doc.PageSetup.TopMargin
    width = page_width - left - doc.PageSetup.RightMargin
    height = page_height - top - doc.PageSetup.BottomMargin

    # 添加矩形形状 (msoShapeRectangle = 1)
    # 参数：Type, Left, Top, Width, Height
    rect = doc.Shapes.AddShape(1, left, top, width, height)

    # 设置矩形样式
    rect.Fill.Visible = 0  # 无填充 (msoFalse)
    rect.Line.Visible = 1  # 显示线条 (msoTrue)
    rect.Line.ForeColor.RGB = 0  # 黑色
    rect.Line.Weight = 1.2  # 线条粗细

    # 在矩形框左上角添加文字
    text_frame = rect.TextFrame
    # 使用制表符分隔左侧和右侧文字
    text_frame.TextRange.Text = " 株机公司普通商密▲5年 \t工艺22"

    # 设置字体：思源宋体，小三（15磅）
    text_frame.TextRange.Font.Name = "思源宋体"
    text_frame.TextRange.Font.Size = 15
    text_frame.TextRange.Font.Color = 0  # 黑色

    # 设置制表位以实现右对齐效果
    # 清除原有制表位
    text_frame.TextRange.ParagraphFormat.TabStops.ClearAll()
    # 添加右对齐制表位
    text_frame.TextRange.ParagraphFormat.TabStops.Add(Position=width - 0.25 * CM_TO_POINT, Alignment=2, Leader=0)

    # 设置对齐方式：左上角
    # msoAnchorTop = 1
    text_frame.VerticalAnchor = 1
    # wdAlignParagraphLeft = 0
    text_frame.TextRange.ParagraphFormat.Alignment = 0

    # # 移除文本框内边距以紧贴边框
    text_frame.MarginLeft = 0
    text_frame.MarginTop = 0
    text_frame.MarginRight = 0
    text_frame.MarginBottom = 0

    rect.ZOrder(5)

    # 在文字下方创建一个新的矩形框（主体内容区域）
    # 预留标题高度约 1.15cm
    header_height = 1.15 * CM_TO_POINT
    # 内部矩形与外部矩形的间距：0.5cm
    padding = 0.5 * CM_TO_POINT

    inner_top = top + header_height
    # 左边距增加 padding
    inner_left = left + padding
    # 宽度减少左右两边的 padding
    inner_width = width - 2 * padding
    # 高度减少顶部的 header_height 和底部的 padding
    inner_height = height - header_height - padding

    # 添加内部矩形
    rect_inner = doc.Shapes.AddShape(1, inner_left, inner_top, inner_width, inner_height)

    # 设置内部矩形样式（与外部一致）
    rect_inner.Fill.Visible = 0
    rect_inner.Line.Visible = 1
    rect_inner.Line.ForeColor.RGB = 0
    rect_inner.Line.Weight = 1.2

    # 在内部矩形框中间上方添加文字
    inner_text_frame = rect_inner.TextFrame
    inner_text_frame.TextRange.Text = "工艺文件"

    # 设置字体：思源宋体，初号（42磅），加粗，黑色
    inner_text_frame.TextRange.Font.Name = "思源宋体"
    inner_text_frame.TextRange.Font.Size = 42
    inner_text_frame.TextRange.Font.Bold = True
    inner_text_frame.TextRange.Font.Color = 0

    # 设置对齐方式：顶端居中
    # msoAnchorTop = 1
    inner_text_frame.VerticalAnchor = 1
    # wdAlignParagraphCenter = 1
    inner_text_frame.TextRange.ParagraphFormat.Alignment = 1

    # 移除文本框内边距
    inner_text_frame.MarginLeft = 0
    inner_text_frame.MarginTop = 0
    inner_text_frame.MarginRight = 0
    inner_text_frame.MarginBottom = 0

    # 在“工艺文件”下方添加产品型号信息
    # 使用独立的文本框，以便于精确定位
    # 位置：内部矩形左上角 + 一定偏移（避开“工艺文件”大字）
    # 假设“工艺文件”高度约 2cm，我们在其下方添加文本框
    model_top_offset = 4.0 * CM_TO_POINT
    model_left_offset = 3.5 * CM_TO_POINT  # 离内部矩形左边界 0.5cm

    # 添加文本框 (使用矩形 msoShapeRectangle = 1 代替 msoShapeTextBox = 17)
    # 某些 Word 版本中 17 对应 msoShapeSmileyFace
    model_textbox = doc.Shapes.AddShape(1, inner_left + model_left_offset, inner_top + model_top_offset, 15 * CM_TO_POINT, 1.5 * CM_TO_POINT)

    # 设置文本框样式：无填充、无边框
    model_textbox.Fill.Visible = 0
    model_textbox.Line.Visible = 0

    # 设置文本内容和格式
    model_frame = model_textbox.TextFrame
    model_range = model_frame.TextRange
    model_range.Text = "产品型号：上海19号线"

    # 全局字体设置：思源宋体，小三（15磅），黑色
    model_range.Font.Name = "思源宋体"
    model_range.Font.Size = 15
    model_range.Font.Color = 0
    model_range.ParagraphFormat.Alignment = 0  # 左对齐
    model_frame.VerticalAnchor = 3 # 垂直居中

    # 对“上海19号线”部分添加下划线
    # 逐个设置字符的下划线属性，避免 Range 计算错误
    # Word 索引从 1 开始
    # "产品型号：" 长度为 5
    # 从第 6 个字符开始设置下划线
    for i in range(6, model_range.Characters.Count + 1):
        model_range.Characters(i).Font.Underline = 6  # wdUnderlineThick (粗下划线)

    # 移除文本框内边距
    model_frame.MarginLeft = 0
    model_frame.MarginTop = 0
    model_frame.MarginRight = 0
    model_frame.MarginBottom = 0
    # 禁用自动换行，防止宽度变化影响文字位置
    model_frame.WordWrap = 0 # msoFalse
    # 清除段落缩进
    model_range.ParagraphFormat.LeftIndent = 0
    model_range.ParagraphFormat.RightIndent = 0

    # 在“工艺文件”下方右侧添加文件名称信息
    # 位置：内部矩形右上角附近
    # 假设文本框宽度 10cm，我们需要计算 Left 坐标使其靠右
    # 这里我们让它距离右边界一定距离，或者直接计算坐标
    # 假设放在右侧，与左侧的“产品型号”在同一水平线上
    name_left_offset = inner_width - 8.0 * CM_TO_POINT - 3.0 * CM_TO_POINT # 离内部矩形右边界 3cm
    
    # 添加文本框 (使用矩形 msoShapeRectangle = 1)
    name_textbox = doc.Shapes.AddShape(1, inner_left + name_left_offset, inner_top + model_top_offset, 15 * CM_TO_POINT, 1.5 * CM_TO_POINT)
    
    # 设置文本框样式：无填充、无边框
    name_textbox.Fill.Visible = 0
    name_textbox.Line.Visible = 0
    
    # 设置文本内容和格式
    name_frame = name_textbox.TextFrame
    name_range = name_frame.TextRange
    name_range.Text = "文件名称：客室座椅安装"
    
    # 全局字体设置：思源宋体，小三（15磅），黑色
    name_range.Font.Name = "思源宋体"
    name_range.Font.Size = 15
    name_range.Font.Color = 0
    name_range.ParagraphFormat.Alignment = 0 # 左对齐 (虽然位置在右侧，但文本框内文字左对齐)
    name_frame.VerticalAnchor = 3 # 垂直居中
    
    # 对“客室座椅安装”部分添加下划线
    # "文件名称：" 长度为 5
    # 从第 6 个字符开始设置下划线
    for i in range(6, name_range.Characters.Count + 1):
        name_range.Characters(i).Font.Underline = 6  # wdUnderlineThick (粗下划线)

    # 移除文本框内边距
    name_frame.MarginLeft = 0
    name_frame.MarginTop = 0
    name_frame.MarginRight = 0
    name_frame.MarginBottom = 0
    # 禁用自动换行
    name_frame.WordWrap = 0
    # 清除段落缩进
    name_range.ParagraphFormat.LeftIndent = 0
    name_range.ParagraphFormat.RightIndent = 0

    # 在“工艺文件”下方左侧添加文件编号信息（第二行）
    # 位置：内部矩形左上角 + 垂直偏移（第一行下方）
    # 假设第一行高度占用约 1.5cm，我们下移一些
    number_top_offset = model_top_offset + 2.0 * CM_TO_POINT # 在“产品型号”下方 1cm
    number_left_offset = model_left_offset # 与“产品型号”左对齐
    
    # 添加文本框 (使用矩形 msoShapeRectangle = 1)
    # 宽度设宽一点以容纳长编号
    number_textbox = doc.Shapes.AddShape(1, inner_left + number_left_offset, inner_top + number_top_offset, 15 * CM_TO_POINT, 1.5 * CM_TO_POINT)
    
    # 设置文本框样式：无填充、无边框
    number_textbox.Fill.Visible = 0
    number_textbox.Line.Visible = 0
    
    # 设置文本内容和格式
    number_frame = number_textbox.TextFrame
    number_range = number_frame.TextRange
    number_range.Text = "文件编号：AJP1023290A-22-01"
    
    # 全局字体设置：思源宋体，小三（15磅），黑色
    number_range.Font.Name = "思源宋体"
    number_range.Font.Size = 15
    number_range.Font.Color = 0
    number_range.ParagraphFormat.Alignment = 0 # 左对齐
    number_frame.VerticalAnchor = 3 # 垂直居中
    
    # 对“AJP1023290A-22-01”部分添加下划线
    # "文件编号：" 长度为 5
    # 从第 6 个字符开始设置下划线
    for i in range(6, number_range.Characters.Count + 1):
        number_range.Characters(i).Font.Underline = 6  # wdUnderlineThick (粗下划线)

    # 移除文本框内边距
    number_frame.MarginLeft = 0
    number_frame.MarginTop = 0
    number_frame.MarginRight = 0
    number_frame.MarginBottom = 0
    # 禁用自动换行
    number_frame.WordWrap = 0
    # 清除段落缩进
    number_range.ParagraphFormat.LeftIndent = 0
    number_range.ParagraphFormat.RightIndent = 0

    # 在“工艺文件”下方右侧添加零部件图号信息（第二行）
    # 位置：内部矩形右上角 + 垂直偏移（第一行下方）
    # 水平位置与第一行的“文件名称”对齐
    part_top_offset = number_top_offset # 与“文件编号”在同一水平线
    part_left_offset = name_left_offset # 与“文件名称”左对齐
    
    # 添加文本框 (使用矩形 msoShapeRectangle = 1)
    # 宽度设宽一点
    part_textbox = doc.Shapes.AddShape(1, inner_left + part_left_offset, inner_top + part_top_offset, 15 * CM_TO_POINT, 1.5 * CM_TO_POINT)
    
    # 设置文本框样式：无填充、无边框
    part_textbox.Fill.Visible = 0
    part_textbox.Line.Visible = 0
    
    # 设置文本内容和格式
    part_frame = part_textbox.TextFrame
    part_range = part_frame.TextRange
    part_range.Text = "零部件图号：AJP1023290A"
    
    # 全局字体设置：思源宋体，小三（15磅），黑色
    part_range.Font.Name = "思源宋体"
    part_range.Font.Size = 15
    part_range.Font.Color = 0
    part_range.ParagraphFormat.Alignment = 0 # 左对齐
    part_frame.VerticalAnchor = 3 # 垂直居中
    
    # 对“AJP1023290A”部分添加下划线
    # "零部件图号：" 长度为 6
    # 从第 7 个字符开始设置下划线
    for i in range(7, part_range.Characters.Count + 1):
        part_range.Characters(i).Font.Underline = 6  # wdUnderlineThick (粗下划线)

    # 移除文本框内边距及其他格式控制
    part_frame.MarginLeft = 0
    part_frame.MarginTop = 0
    part_frame.MarginRight = 0
    part_frame.MarginBottom = 0
    part_frame.WordWrap = 0
    part_range.ParagraphFormat.LeftIndent = 0
    part_range.ParagraphFormat.RightIndent = 0

    # 在下方添加编制、校对、审核信息（第三行）
    # 位置：内部矩形水平居中，垂直位于第二行下方
    sign_top_offset = part_top_offset + 5.0 * CM_TO_POINT # 在第二行下方 1.2cm
    
    # 文本框宽度设为内部矩形宽度，以便居中
    sign_width = inner_width
    sign_height = 1.5 * CM_TO_POINT
    
    # 添加文本框 (使用矩形 msoShapeRectangle = 1)
    sign_textbox = doc.Shapes.AddShape(1, inner_left + 4.0 * CM_TO_POINT, inner_top + sign_top_offset, sign_width, sign_height)
    
    # 设置文本框样式：无填充、无边框
    sign_textbox.Fill.Visible = 0
    sign_textbox.Line.Visible = 0
    
    # 设置文本内容
    sign_frame = sign_textbox.TextFrame
    sign_range = sign_frame.TextRange
    # 注意：文本中包含中文冒号和空格
    text_content = "编制 ：黎运阳  2026-01-21；校对：张权  2026-01-21；审核：毛幸福  2026-01-21；"
    sign_range.Text = text_content
    
    # 全局字体设置：思源宋体，小三（15磅），黑色
    sign_range.Font.Name = "思源宋体"
    sign_range.Font.Size = 13
    sign_range.Font.Color = 0
    sign_range.ParagraphFormat.Alignment = 1 # 居中对齐
    sign_frame.VerticalAnchor = 3 # 垂直居中
    
    # 对人名和时间部分添加下划线
    # 需要识别具体的字符范围
    # "编制 ：" (4字符) -> "黎运阳  2026-01-21" (下划线) -> "；校对：" (4字符) ...
    # 为了简化，我们使用查找方式或硬编码位置（如果文本固定）
    # 这里文本是固定的，我们可以直接计算位置
    # 编制 ：(4) -> 黎运阳  2026-01-21 (15) -> ；校对：(4) -> 张权  2026-01-21 (14) -> ；审核：(4) -> 毛幸福  2026-01-21 (15) -> ；(1)
    
    # 定义下划线片段的起始位置和长度 (Word索引从1开始)
    # 片段1: "黎运阳  2026-01-21"
    # start = 1 + 4 = 5
    # length = 15 (黎运阳 3 + 空格 2 + 日期 10) -> 实际上是 "黎运阳  2026-01-21"
    # 片段2: "张权  2026-01-21"
    # start = 5 + 15 + 4 = 24
    # length = 14 (张权 2 + 空格 2 + 日期 10)
    # 片段3: "毛幸福  2026-01-21"
    # start = 24 + 14 + 4 = 42
    # length = 15 (毛幸福 3 + 空格 2 + 日期 10)
    
    # 为防止计算错误，我们使用 find 功能或逐段设置
    # 这里使用循环范围设置
    
    segments = [
        (5, 15),  # 黎运阳  2026-01-21
        (24, 14), # 张权  2026-01-21
        (42, 15)  # 毛幸福  2026-01-21
    ]
    
    for start, length in segments:
        for i in range(start, start + length):
            if i <= sign_range.Characters.Count:
                sign_range.Characters(i).Font.Underline = 6 # wdUnderlineThick
    
    # 移除文本框内边距及其他格式控制
    sign_frame.MarginLeft = 0
    sign_frame.MarginTop = 0
    sign_frame.MarginRight = 0
    sign_frame.MarginBottom = 0
    sign_frame.WordWrap = 0
    sign_range.ParagraphFormat.LeftIndent = 0
    sign_range.ParagraphFormat.RightIndent = 0

    # 在编制校对审核信息下方添加标准化、会签、批准信息（第四行）
    # 位置：内部矩形水平居中，垂直位于第三行下方
    approve_top_offset = sign_top_offset + 1.2 * CM_TO_POINT # 在编制行下方 1.2cm
    
    # 添加文本框 (使用矩形 msoShapeRectangle = 1)
    approve_textbox = doc.Shapes.AddShape(1, inner_left + 4.0 * CM_TO_POINT, inner_top + approve_top_offset, sign_width, sign_height)
    
    # 设置文本框样式：无填充、无边框
    approve_textbox.Fill.Visible = 0
    approve_textbox.Line.Visible = 0
    
    # 设置文本内容
    approve_frame = approve_textbox.TextFrame
    approve_range = approve_frame.TextRange
    text_content_2 = "标准化：黎运阳  2026-01-21；会签：张权  2026-01-21；批准：毛幸福  2026-01-21；"
    approve_range.Text = text_content_2
    
    # 全局字体设置：思源宋体，小三（15磅），黑色
    approve_range.Font.Name = "思源宋体"
    approve_range.Font.Size = 13
    approve_range.Font.Color = 0
    approve_range.ParagraphFormat.Alignment = 1 # 居中对齐
    approve_frame.VerticalAnchor = 3 # 垂直居中
    
    # 对人名和时间部分添加下划线
    # 文本结构与上一行类似
    # 标准化：(4) -> 黎运阳  2026-01-21 (15) -> ；会签：(4) -> 张权  2026-01-21 (14) -> ；批准：(4) -> 毛幸福  2026-01-21 (15) -> ；(1)
    # 起始位置与上一行略有不同，因为“标准化”是3个字，但冒号前缀长度实际上需要重新计算
    # "标准化：" (4字符)
    # 片段1: "黎运阳  2026-01-21" -> start = 1 + 4 = 5, length = 15
    # "；会签：" (4字符)
    # 片段2: "张权  2026-01-21" -> start = 5 + 15 + 4 = 24, length = 14
    # "；批准：" (4字符)
    # 片段3: "毛幸福  2026-01-21" -> start = 24 + 14 + 4 = 42, length = 15
    
    # 看起来结构长度一致
    segments_2 = [
        (5, 15),  # 黎运阳  2026-01-21
        (24, 14), # 张权  2026-01-21
        (42, 15)  # 毛幸福  2026-01-21
    ]
    
    for start, length in segments_2:
        for i in range(start, start + length):
            if i <= approve_range.Characters.Count:
                approve_range.Characters(i).Font.Underline = 6 # wdUnderlineThick

    # 移除文本框内边距及其他格式控制
    approve_frame.MarginLeft = 0
    approve_frame.MarginTop = 0
    approve_frame.MarginRight = 0
    approve_frame.MarginBottom = 0
    approve_frame.WordWrap = 0
    approve_range.ParagraphFormat.LeftIndent = 0
    approve_range.ParagraphFormat.RightIndent = 0

    # 在批准信息下方添加公司名称（第五行）
    # 位置：内部矩形水平居中，垂直位于第四行下方
    company_top_offset = approve_top_offset + 3.5 * CM_TO_POINT # 在批准行下方 1.2cm
    
    # 添加文本框 (使用矩形 msoShapeRectangle = 1)
    company_textbox = doc.Shapes.AddShape(1, inner_left + 8.0 * CM_TO_POINT, inner_top + company_top_offset, sign_width, sign_height)
    
    # 设置文本框样式：无填充、无边框
    company_textbox.Fill.Visible = 0
    company_textbox.Line.Visible = 0
    
    # 设置文本内容
    company_frame = company_textbox.TextFrame
    company_range = company_frame.TextRange
    company_range.Text = "中车株洲电力机车有限公司城轨制造中心"
    
    # 全局字体设置：思源宋体，小三（15磅），黑色，加粗
    company_range.Font.Name = "思源宋体"
    company_range.Font.Size = 15
    company_range.Font.Color = 0
    company_range.Font.Bold = True # 加粗
    company_range.ParagraphFormat.Alignment = 1 # 居中对齐
    company_frame.VerticalAnchor = 3 # 垂直居中

    # 移除文本框内边距及其他格式控制
    company_frame.MarginLeft = 0
    company_frame.MarginTop = 0
    company_frame.MarginRight = 0
    company_frame.MarginBottom = 0
    company_frame.WordWrap = 0
    company_range.ParagraphFormat.LeftIndent = 0
    company_range.ParagraphFormat.RightIndent = 0

    # 在公司名称下方添加日期页码信息（第六行）
    # 位置：内部矩形水平居中，垂直位于第五行下方
    date_top_offset = company_top_offset + 1.2 * CM_TO_POINT # 在公司名称行下方 1.2cm
    
    # 添加文本框 (使用矩形 msoShapeRectangle = 1)
    date_textbox = doc.Shapes.AddShape(1, inner_left  + 8.5 * CM_TO_POINT, inner_top + date_top_offset, sign_width, sign_height)
    
    # 设置文本框样式：无填充、无边框
    date_textbox.Fill.Visible = 0
    date_textbox.Line.Visible = 0
    
    # 设置文本内容
    date_frame = date_textbox.TextFrame
    date_range = date_frame.TextRange
    date_range.Text = "2026年01月21日第1版共8页"
    
    # 全局字体设置：思源宋体，小三（15磅），黑色
    date_range.Font.Name = "思源宋体"
    date_range.Font.Size = 15
    date_range.Font.Color = 0
    # 注意：不需要加粗
    date_range.ParagraphFormat.Alignment = 1 # 居中对齐
    date_frame.VerticalAnchor = 3 # 垂直居中

    # 移除文本框内边距及其他格式控制
    date_frame.MarginLeft = 0
    date_frame.MarginTop = 0
    date_frame.MarginRight = 0
    date_frame.MarginBottom = 0
    date_frame.WordWrap = 0
    date_range.ParagraphFormat.LeftIndent = 0
    date_range.ParagraphFormat.RightIndent = 0

    # 移动光标到文档末尾，确保在最后插入分页符
    word.Selection.EndKey(Unit=6) # wdStory
    
    # 插入分页符，创建新页面
    # wdPageBreak = 7
    word.Selection.InsertBreak(Type=7)
    
    # 在新页面（第二页）绘制外边框矩形
    # 注意：Shapes.AddShape 默认锚定到当前页或当前 Selection
    # 为了确保添加在第二页，我们需要指定锚点或者确保 Selection 在第二页
    # 前面已经将 Selection 移动到了文档末尾（第二页）
    
    # 再次移动光标到文档末尾，准备在新页面添加内容
    # 这一步至关重要，因为 InsertBreak 后光标可能还在前一页末尾或分页符处
    word.Selection.EndKey(Unit=6) # wdStory

    # 添加第二页的外矩形
    # 只要光标在第二页，Shapes.AddShape 就会默认锚定到第二页
    # 参数与第一页一致
    rect_page2 = doc.Shapes.AddShape(1, left, top, width, height)
    
    # 设置样式
    rect_page2.Fill.Visible = 0
    rect_page2.Line.Visible = 1
    rect_page2.Line.ForeColor.RGB = 0
    rect_page2.Line.Weight = 1.2

   # 在矩形框左上角添加文字
    text_frame_2 = rect_page2.TextFrame
    # 使用制表符分隔左侧和右侧文字
    text_frame_2.TextRange.Text = " 株机公司普通商密▲5年 \t工艺22"

    # 设置字体：思源宋体，小三（15磅）
    text_frame_2.TextRange.Font.Name = "思源宋体"
    text_frame_2.TextRange.Font.Size = 15
    text_frame_2.TextRange.Font.Color = 0  # 黑色

    # 设置制表位以实现右对齐效果
    # 清除原有制表位
    text_frame_2.TextRange.ParagraphFormat.TabStops.ClearAll()
    # 添加右对齐制表位
    text_frame_2.TextRange.ParagraphFormat.TabStops.Add(Position=width - 0.25 * CM_TO_POINT, Alignment=2, Leader=0)

    # 设置对齐方式：左上角
    # msoAnchorTop = 1
    text_frame_2.VerticalAnchor = 1
    # wdAlignParagraphLeft = 0
    text_frame_2.TextRange.ParagraphFormat.Alignment = 0

    # # 移除文本框内边距以紧贴边框
    text_frame_2.MarginLeft = 0
    text_frame_2.MarginTop = 0
    text_frame_2.MarginRight = 0
    text_frame_2.MarginBottom = 0

    # 将文本框设置为衬于文字下方
    # msoSendBehindText = 5
    rect_page2.ZOrder(5)

    # 在文字下方创建一个新的矩形框（主体内容区域）
    # 预留标题高度约 1.15cm
    header_height = 1.15 * CM_TO_POINT
    # 内部矩形与外部矩形的间距：0.5cm
    padding = 0.5 * CM_TO_POINT

    inner_top = top + header_height
    # 左边距增加 padding
    inner_left = left + padding
    # 宽度减少左右两边的 padding
    inner_width = width - 2 * padding
    # 高度减少顶部的 header_height 和底部的 padding
    inner_height = height - header_height - padding

    # 添加内部矩形
    rect_inner2 = doc.Shapes.AddShape(1, inner_left, inner_top, inner_width, inner_height)

    # 设置内部矩形样式（与外部一致）
    rect_inner2.Fill.Visible = 0
    rect_inner2.Line.Visible = 1
    rect_inner2.Line.ForeColor.RGB = 0
    rect_inner2.Line.Weight = 1.2

    # 在内部矩形框中间上方添加文字
    inner_text_frame_2 = rect_inner2.TextFrame
    inner_text_frame_2.TextRange.Text = "文件变更记录卡"
    
    # 设置字体：思源宋体，二号（22磅），加粗，黑色
    inner_text_frame_2.TextRange.Font.Name = "思源宋体"
    inner_text_frame_2.TextRange.Font.Size = 22
    inner_text_frame_2.TextRange.Font.Bold = True
    inner_text_frame_2.TextRange.Font.Color = 0
    
    # 设置对齐方式：顶端居中
    inner_text_frame_2.VerticalAnchor = 1 # msoAnchorTop
    inner_text_frame_2.TextRange.ParagraphFormat.Alignment = 1 # wdAlignParagraphCenter
    
    # 移除文本框内边距
    inner_text_frame_2.MarginLeft = 0
    inner_text_frame_2.MarginTop = 0
    inner_text_frame_2.MarginRight = 0
    inner_text_frame_2.MarginBottom = 0

    # 将文本框设置为衬于文字下方
    # msoSendBehindText = 5
    rect_inner2.ZOrder(5)

    # 在“文件变更记录卡”下方添加表格
    # 表格位置：内部矩形内，距离左右各 0.1cm
    table_padding = 0.5 * CM_TO_POINT
    table_left = inner_left + table_padding
    table_width = inner_width - 2 * table_padding
    
    # 垂直位置：标题文字大约占用 1cm - 1.5cm，我们在其下方添加表格
    # 假设标题文字区域高度 1.5cm
    table_top_offset = 3 * CM_TO_POINT
    # 表格的垂直位置不能直接通过 Shapes.AddTable 指定绝对坐标（AddTable 返回 Table 对象，通常插入到 Range）
    # 但我们可以使用 doc.Tables.Add(Range, NumRows, NumColumns)
    # 为了精确定位，最好的方法是将表格放在一个文本框内，或者使用 Shapes.AddTable（如果版本支持且也是 Shape）
    # 或者直接在页面上添加表格，然后设置其 WrapFormat 和位置
    
    # 这里我们尝试直接在文档中添加表格，并设置其位置为固定
    # 首先需要一个 Range，我们创建一个位于文档末尾的 Range（第二页）
    # 注意：前面的 Selection 已经在第二页末尾
    
    # 添加表格：5行2列
    # 使用 doc.Tables.Add 方法
    table = doc.Tables.Add(word.Selection.Range, 5, 2)
    
    # 设置表格属性
    table.PreferredWidthType = 3 # wdPreferredWidthPoints
    table.PreferredWidth = table_width
    
    # 设置表格位置（浮动表格）
    table.Rows.WrapAroundText = 1 # True，使其成为浮动表格以便定位
    # 设置绝对位置
    table.Rows.HorizontalPosition = table_left
    table.Rows.VerticalPosition = inner_top + table_top_offset
    # 相对位置参考：页面
    table.Rows.RelativeHorizontalPosition = 1 # wdRelativeHorizontalPositionPage
    table.Rows.RelativeVerticalPosition = 1 # wdRelativeVerticalPositionPage
    
    # 设置表格边框
    table.Borders.Enable = 1 # 启用所有边框
    
    # 设置列宽
    # 第一列宽度：3cm (约 85磅)
    col1_width = 6.0 * CM_TO_POINT
    table.Columns(1).Width = col1_width
    # 第二列宽度：剩余宽度
    table.Columns(2).Width = table_width - col1_width

    # 填充第一列内容并设置格式
    row_titles = ["文件编号", "文件名称", "产品图号", "项目名称", "车种/工区工位号"]
    
    for i, title in enumerate(row_titles):
        # 获取单元格 (行索引从 1 开始)
        cell = table.Cell(i + 1, 1)
        # 设置文本
        # 注意：Cell.Range.Text 赋值会自动包含结束符，所以直接赋值即可
        cell.Range.Text = title
        
        # 设置单元格垂直居中
        # wdCellAlignVerticalCenter = 1
        cell.VerticalAlignment = 1
        
        # 设置段落格式（水平居中）
        # wdAlignParagraphCenter = 1
        para_format = cell.Range.ParagraphFormat
        para_format.Alignment = 1
        
        # 设置行距和段间距
        # 单倍行距
        para_format.LineSpacingRule = 0 # wdLineSpaceSingle
        # 段前段后为0
        para_format.SpaceBefore = 0
        para_format.SpaceAfter = 0
        # 不对齐到网格
        para_format.DisableLineHeightGrid = True
        
        # 设置字体格式
        # 思源黑体，小三（15磅）
        cell.Range.Font.Name = "思源黑体"
        cell.Range.Font.Size = 15
        cell.Range.Font.Color = 0 # 黑色

    # 在表格正下方添加“文件版本历史记录”文本
    # 计算位置：表格顶部位置 + 表格高度 + 间距
    # 假设表格高度：5行 * 每行高度（Word 默认行高或根据内容自适应）
    # 这里我们估算一下或者设置一个固定的偏移量
    # 假设表格总高度约 5cm (每行 1cm)
    history_title_top = inner_top + table_top_offset + 1.0 * CM_TO_POINT
    
    # 添加文本框 (使用矩形 msoShapeRectangle = 1)
    history_textbox = doc.Shapes.AddShape(1, -table_padding, history_title_top, inner_width, 1.5 * CM_TO_POINT)
    
    # 设置文本框样式：无填充、无边框
    history_textbox.Fill.Visible = 0
    history_textbox.Line.Visible = 0
    
    # 设置文本内容
    history_frame = history_textbox.TextFrame
    history_range = history_frame.TextRange
    history_range.Text = "文件版本历史记录"
    
    # 设置字体：思源宋体，二号（22磅），加粗，黑色
    history_range.Font.Name = "思源宋体"
    history_range.Font.Size = 22
    history_range.Font.Bold = True
    history_range.Font.Color = 0
    
    # 设置对齐方式：居中
    history_range.ParagraphFormat.Alignment = 1 # 居中对齐
    history_frame.VerticalAnchor = 1 # 顶端对齐 (或者居中)
    
    # 移除文本框内边距
    history_frame.MarginLeft = 0
    history_frame.MarginTop = 0
    history_frame.MarginRight = 0
    history_frame.MarginBottom = 0
    
    # 设置衬于文字下方
    history_textbox.ZOrder(5)

    # 在“文件版本历史记录”下方添加第二个表格
    # 4行5列
    # 垂直位置：在标题下方 3.5cm 处，增加间距
    # 注意：history_title_top 是标题的 Top 坐标
    # 我们希望表格 Top = history_title_top + 间距
    # 但 VerticalPosition 是相对于 Page 的（如果 RelativeVerticalPosition=1）
    # 这里我们直接设置绝对位置
    table2_vertical_pos = history_title_top + 7.0 * CM_TO_POINT
    
    # 将光标移动到文档末尾，跳出上一个表格范围
    word.Selection.EndKey(Unit=6) # wdStory
    
    # 添加表格：4行5列
    # 使用 doc.Tables.Add 方法
    table2 = doc.Tables.Add(word.Selection.Range, 4, 5)
    
    # 设置表格属性
    table2.PreferredWidthType = 3 # wdPreferredWidthPoints
    table2.PreferredWidth = table_width # 复用之前的表格宽度
    
    # 设置表格位置（浮动表格）
    table2.Rows.WrapAroundText = 1 # True
    table2.Rows.HorizontalPosition = table_left # 复用之前的左边距
    table2.Rows.VerticalPosition = table2_vertical_pos
    table2.Rows.RelativeHorizontalPosition = 1 # wdRelativeHorizontalPositionPage
    table2.Rows.RelativeVerticalPosition = 1 # wdRelativeVerticalPositionPage
    
    # 设置表格边框
    table2.Borders.Enable = 1

    # 填充第二页表格第一行表头
    table2_headers = ["版本号", "实施日期", "编制者", "变更记录", "文件状态"]
    for i, header in enumerate(table2_headers):
        cell = table2.Cell(1, i + 1)
        cell.Range.Text = header
        
        # 设置单元格样式
        cell.VerticalAlignment = 1 # wdCellAlignVerticalCenter
        
        para_format = cell.Range.ParagraphFormat
        para_format.Alignment = 1 # wdAlignParagraphCenter
        para_format.LineSpacingRule = 0 # wdLineSpaceSingle
        para_format.SpaceAfter = 0
        para_format.DisableLineHeightGrid = True # 不对齐文档网格
        
        cell.Range.Font.Name = "思源黑体"
        cell.Range.Font.Size = 15
        cell.Range.Font.Color = 0 # 黑色

    # 移动光标到文档末尾，跳出上一个表格范围
    word.Selection.EndKey(Unit=6) # wdStory
    
    # 插入分页符，创建新页面（第三页）
    # wdPageBreak = 7
    word.Selection.InsertBreak(Type=7)
    
    # 再次移动光标到文档末尾，准备在新页面添加内容
    word.Selection.EndKey(Unit=6) # wdStory

    # 添加第三页的外矩形
    # 只要光标在第三页，Shapes.AddShape 就会默认锚定到第三页
    # 参数与第一页一致
    rect_page3 = doc.Shapes.AddShape(1, left, top, width, height)
    
    # 设置样式
    rect_page3.Fill.Visible = 0
    rect_page3.Line.Visible = 1
    rect_page3.Line.ForeColor.RGB = 0
    rect_page3.Line.Weight = 1.2

    # 将文本框设置为衬于文字下方
    # msoSendBehindText = 5
    rect_page3.ZOrder(5)

    # 在第三页框内部创建表格：6行16列
    # 表格紧贴框线内部
    # 顶部位置：top + header_height (预留标题高度，与前几页一致)
    table3_width = inner_width
    # 高度占据剩余空间
    table3_height = inner_height

    # 将光标移动到文档末尾
    word.Selection.EndKey(Unit=6) # wdStory

    # 添加表格：2行7列
    table3 = doc.Tables.Add(word.Selection.Range, 2, 7)

    # 设置表格属性
    table3.PreferredWidthType = 3 # wdPreferredWidthPoints

    # 设置表格位置（浮动表格）
    table3.Rows.WrapAroundText = 1 # True
    table3.Rows.HorizontalPosition = 2.2 * CM_TO_POINT
    table3.Rows.VerticalPosition = 0.5 * CM_TO_POINT
    table3.Rows.RelativeHorizontalPosition = 1 # wdRelativeHorizontalPositionPage
    table3.Rows.RelativeVerticalPosition = 1 # wdRelativeVerticalPositionPage

    # 设置表格边框
    table3.Borders.Enable = 1

    # 填充表格内容
    # 单元格1行1列写入“中车株洲电力机车有限公司”
    cell_1_1 = table3.Cell(1, 1)
    cell_1_1.Range.Text = "中车株洲电力机车有限公司"
    cell_1_1.VerticalAlignment = 1 # 垂直居中
    cell_1_1.Range.ParagraphFormat.Alignment = 1 # 水平居中
    cell_1_1.Range.ParagraphFormat.LineSpacingRule = 0 # 单倍行距
    cell_1_1.Range.ParagraphFormat.SpaceAfter = 0
    cell_1_1.Range.ParagraphFormat.DisableLineHeightGrid = True
    cell_1_1.Range.Font.Name = "思源黑体"
    cell_1_1.Range.Font.Size = 11
    cell_1_1.Range.Font.Color = 0

    # 单元格2行1列写入“城轨制造中心”
    cell_2_1 = table3.Cell(2, 1)
    cell_2_1.Range.Text = "城轨制造中心"
    cell_2_1.VerticalAlignment = 1 # 垂直居中
    cell_2_1.Range.ParagraphFormat.Alignment = 1 # 水平居中
    cell_2_1.Range.ParagraphFormat.LineSpacingRule = 0 # 单倍行距
    cell_2_1.Range.ParagraphFormat.SpaceAfter = 0
    cell_2_1.Range.ParagraphFormat.DisableLineHeightGrid = True
    cell_2_1.Range.Font.Name = "思源黑体"
    cell_2_1.Range.Font.Size = 11
    cell_2_1.Range.Font.Color = 0
    
    # 单元格1行2列和2行2列合并为一个单元格写入“组装工序卡”
    cell_1_2 = table3.Cell(1, 2)
    cell_2_2 = table3.Cell(2, 2)
    cell_1_2.Merge(cell_2_2)
    # 合并后使用 cell_1_2 访问
    cell_1_2.Range.Text = "组装工序卡"
    cell_1_2.VerticalAlignment = 1 # 垂直居中
    cell_1_2.Range.ParagraphFormat.Alignment = 1 # 水平居中
    cell_1_2.Range.ParagraphFormat.LineSpacingRule = 0 # 单倍行距
    cell_1_2.Range.ParagraphFormat.SpaceAfter = 0
    cell_1_2.Range.ParagraphFormat.DisableLineHeightGrid = True
    cell_1_2.Range.Font.Name = "思源黑体"
    cell_1_2.Range.Font.Size = 15
    cell_1_2.Range.Font.Color = 0
    
    # 给“组装”两个字添加下划线
    # "组装工序卡" -> 前两个字符
    # Word Range 索引从 0 开始 (Python 切片风格) 或者 Characters 集合从 1 开始
    # 使用 Characters(Start, End) 或者 Range(Start, End)
    # 注意：Cell.Range 包含单元格结束符，所以需要小心处理
    # 直接获取 Characters(1) 和 Characters(2)
    cell_1_2.Range.Characters(1).Font.Underline = 1 # wdUnderlineSingle
    cell_1_2.Range.Characters(2).Font.Underline = 1 # wdUnderlineSingle
    
    # 单元格1行3列写入“文件名称”
    cell_1_3 = table3.Cell(1, 3)
    cell_1_3.Range.Text = "文件名称"
    cell_1_3.VerticalAlignment = 1 # 垂直居中
    cell_1_3.Range.ParagraphFormat.Alignment = 1 # 水平居中
    cell_1_3.Range.ParagraphFormat.LineSpacingRule = 0 # 单倍行距
    cell_1_3.Range.ParagraphFormat.SpaceAfter = 0
    cell_1_3.Range.ParagraphFormat.DisableLineHeightGrid = True
    cell_1_3.Range.Font.Name = "思源黑体"
    cell_1_3.Range.Font.Size = 11
    cell_1_3.Range.Font.Color = 0
    
    # 单元格2行3列写入“内装工位”
    cell_2_3 = table3.Cell(2, 3)
    cell_2_3.Range.Text = "内装工位"
    cell_2_3.VerticalAlignment = 1 # 垂直居中
    cell_2_3.Range.ParagraphFormat.Alignment = 1 # 水平居中
    cell_2_3.Range.ParagraphFormat.LineSpacingRule = 0 # 单倍行距
    cell_2_3.Range.ParagraphFormat.SpaceAfter = 0
    cell_2_3.Range.ParagraphFormat.DisableLineHeightGrid = True
    cell_2_3.Range.Font.Name = "思源黑体"
    cell_2_3.Range.Font.Size = 11
    cell_2_3.Range.Font.Color = 0
    
    # 单元格1行4列写入“工序标识”
    cell_1_4 = table3.Cell(1, 4)
    cell_1_4.Range.Text = "工序标识"
    cell_1_4.VerticalAlignment = 1 # 垂直居中
    cell_1_4.Range.ParagraphFormat.Alignment = 1 # 水平居中
    cell_1_4.Range.ParagraphFormat.LineSpacingRule = 0 # 单倍行距
    cell_1_4.Range.ParagraphFormat.SpaceAfter = 0
    cell_1_4.Range.ParagraphFormat.DisableLineHeightGrid = True
    cell_1_4.Range.Font.Name = "思源黑体"
    cell_1_4.Range.Font.Size = 11
    cell_1_4.Range.Font.Color = 0
    
    # 单元格2行4列写入“5214853”
    cell_2_4 = table3.Cell(2, 4)
    cell_2_4.Range.Text = "5214853"
    cell_2_4.VerticalAlignment = 1 # 垂直居中
    cell_2_4.Range.ParagraphFormat.Alignment = 1 # 水平居中
    cell_2_4.Range.ParagraphFormat.LineSpacingRule = 0 # 单倍行距
    cell_2_4.Range.ParagraphFormat.SpaceAfter = 0
    cell_2_4.Range.ParagraphFormat.DisableLineHeightGrid = True
    cell_2_4.Range.Font.Name = "思源黑体"
    cell_2_4.Range.Font.Size = 11
    cell_2_4.Range.Font.Color = 0
    table3.Cell(2, 4).Range.Text = "5214853"
    
    # 单元格1行5列写入“图号”
    cell_1_5 = table3.Cell(1, 5)
    cell_1_5.Range.Text = "图号"
    cell_1_5.VerticalAlignment = 1 # 垂直居中
    cell_1_5.Range.ParagraphFormat.Alignment = 1 # 水平居中
    cell_1_5.Range.ParagraphFormat.LineSpacingRule = 0 # 单倍行距
    cell_1_5.Range.ParagraphFormat.SpaceAfter = 0
    cell_1_5.Range.ParagraphFormat.DisableLineHeightGrid = True
    cell_1_5.Range.Font.Name = "思源黑体"
    cell_1_5.Range.Font.Size = 11
    cell_1_5.Range.Font.Color = 0
    
    
    # 单元格2行5列写入“AJP1023290A”
    cell_2_5 = table3.Cell(2, 5)
    cell_2_5.Range.Text = "AJP1023290A"
    cell_2_5.VerticalAlignment = 1 # 垂直居中
    cell_2_5.Range.ParagraphFormat.Alignment = 1 # 水平居中
    cell_2_5.Range.ParagraphFormat.LineSpacingRule = 0 # 单倍行距
    cell_2_5.Range.ParagraphFormat.SpaceAfter = 0
    cell_2_5.Range.ParagraphFormat.DisableLineHeightGrid = True
    cell_2_5.Range.Font.Name = "思源黑体"
    cell_2_5.Range.Font.Size = 11
    cell_2_5.Range.Font.Color = 0
    
    # 单元格1行7列写入“工序工时(min)”
    cell_1_7 = table3.Cell(1, 7)
    cell_1_7.Range.Text = "工序工时(min)"
    cell_1_7.VerticalAlignment = 1 # 垂直居中
    cell_1_7.Range.ParagraphFormat.Alignment = 1 # 水平居中
    cell_1_7.Range.ParagraphFormat.LineSpacingRule = 0 # 单倍行距
    cell_1_7.Range.ParagraphFormat.SpaceAfter = 0
    cell_1_7.Range.ParagraphFormat.DisableLineHeightGrid = True
    cell_1_7.Range.Font.Name = "思源黑体"
    cell_1_7.Range.Font.Size = 11
    cell_1_7.Range.Font.Color = 0
    
    # 单元格2行7列写入“650min”
    cell_2_7 = table3.Cell(2, 7)
    cell_2_7.Range.Text = "650min"
    cell_2_7.VerticalAlignment = 1 # 垂直居中
    cell_2_7.Range.ParagraphFormat.Alignment = 1 # 水平居中
    cell_2_7.Range.ParagraphFormat.LineSpacingRule = 0 # 单倍行距
    cell_2_7.Range.ParagraphFormat.SpaceAfter = 0
    cell_2_7.Range.ParagraphFormat.DisableLineHeightGrid = True
    cell_2_7.Range.Font.Name = "思源黑体"
    cell_2_7.Range.Font.Size = 11
    cell_2_7.Range.Font.Color = 0


    # # 统一设置格式
    # for row in table3.Rows:
    #     for cell in row.Cells:
    #         cell.VerticalAlignment = 1 # 垂直居中
    #         cell.Range.ParagraphFormat.Alignment = 1 # 水平居中
    #         cell.Range.ParagraphFormat.LineSpacingRule = 0 # 单倍行距
    #         cell.Range.ParagraphFormat.SpaceAfter = 0
    #         cell.Range.ParagraphFormat.DisableLineHeightGrid = True
    #         cell.Range.Font.Name = "思源黑体"
    #         cell.Range.Font.Size = 15
    #         cell.Range.Font.Color = 0

    # 保存文档
    # FileFormat=12 代表 docx 格式 (wdFormatXMLDocument)
    doc.SaveAs(str(file_path), FileFormat=12)
    # 关闭文档和 Word 应用程序
    # doc.Close()
    # word.Quit()


if __name__ == '__main__':
    create_document(file_path=Path(__file__).parent.parent / 'source' / 'test.docx', item={})
