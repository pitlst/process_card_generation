import win32com.client
from pathlib import Path


def create_document(file_path: Path, item: dict):
    # 获取绝对路径，Word COM 接口通常需要绝对路径
    file_path = file_path.absolute()
    # 检查文件是否已存在，如果存在先删除，避免 SaveAs 弹窗或报错
    if file_path.exists():
        file_path.unlink()
    # 启动 Word 应用程序
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # 设置为 True 可以看到 Word 界面，方便调试；设置为 False 则后台运行
    word.DisplayAlerts = 0  # 禁用警告弹窗

    # 新建文档
    doc = word.Documents.Add()

    # 设置页面布局：A4 横向
    # wdOrientLandscape = 1, wdPaperA4 = 7
    doc.PageSetup.Orientation = 1
    doc.PageSetup.PaperSize = 7

    # 设置页边距：左 2.2cm，上下右 0.5cm
    # 直接计算点数 (1 cm ≈ 28.35 points)，避免调用 word.CentimetersToPoints 可能出现的 COM 错误
    cm_to_points = 28.35
    doc.PageSetup.LeftMargin = 2.2 * cm_to_points
    doc.PageSetup.RightMargin = 0.5 * cm_to_points
    doc.PageSetup.TopMargin = 0.5 * cm_to_points
    doc.PageSetup.BottomMargin = 0.5 * cm_to_points

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
    text_frame.TextRange.ParagraphFormat.TabStops.Add(Position=width - 0.25 * cm_to_points, Alignment=2, Leader=0)

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

    # 在文字下方创建一个新的矩形框（主体内容区域）
    # 预留标题高度约 1.15cm
    header_height = 1.15 * cm_to_points
    # 内部矩形与外部矩形的间距：0.5cm
    padding = 0.5 * cm_to_points

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
    model_top_offset = 4.0 * cm_to_points
    model_left_offset = 3.5 * cm_to_points  # 离内部矩形左边界 0.5cm

    # 添加文本框 (使用矩形 msoShapeRectangle = 1 代替 msoShapeTextBox = 17)
    # 某些 Word 版本中 17 对应 msoShapeSmileyFace
    model_textbox = doc.Shapes.AddShape(1, inner_left + model_left_offset, inner_top + model_top_offset, 15 * cm_to_points, 1.5 * cm_to_points)

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
    name_left_offset = inner_width - 8.0 * cm_to_points - 3.0 * cm_to_points # 离内部矩形右边界 3cm
    
    # 添加文本框 (使用矩形 msoShapeRectangle = 1)
    name_textbox = doc.Shapes.AddShape(1, inner_left + name_left_offset, inner_top + model_top_offset, 15 * cm_to_points, 1.5 * cm_to_points)
    
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
    number_top_offset = model_top_offset + 2.0 * cm_to_points # 在“产品型号”下方 1cm
    number_left_offset = model_left_offset # 与“产品型号”左对齐
    
    # 添加文本框 (使用矩形 msoShapeRectangle = 1)
    # 宽度设宽一点以容纳长编号
    number_textbox = doc.Shapes.AddShape(1, inner_left + number_left_offset, inner_top + number_top_offset, 15 * cm_to_points, 1.5 * cm_to_points)
    
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
    part_textbox = doc.Shapes.AddShape(1, inner_left + part_left_offset, inner_top + part_top_offset, 15 * cm_to_points, 1.5 * cm_to_points)
    
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
    sign_top_offset = part_top_offset + 5.0 * cm_to_points # 在第二行下方 1.2cm
    
    # 文本框宽度设为内部矩形宽度，以便居中
    sign_width = inner_width
    sign_height = 1.5 * cm_to_points
    
    # 添加文本框 (使用矩形 msoShapeRectangle = 1)
    sign_textbox = doc.Shapes.AddShape(1, inner_left + 4.0 * cm_to_points, inner_top + sign_top_offset, sign_width, sign_height)
    
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
    approve_top_offset = sign_top_offset + 1.2 * cm_to_points # 在编制行下方 1.2cm
    
    # 添加文本框 (使用矩形 msoShapeRectangle = 1)
    approve_textbox = doc.Shapes.AddShape(1, inner_left + 4.0 * cm_to_points, inner_top + approve_top_offset, sign_width, sign_height)
    
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
    company_top_offset = approve_top_offset + 3.5 * cm_to_points # 在批准行下方 1.2cm
    
    # 添加文本框 (使用矩形 msoShapeRectangle = 1)
    company_textbox = doc.Shapes.AddShape(1, inner_left + 8.0 * cm_to_points, inner_top + company_top_offset, sign_width, sign_height)
    
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
    date_top_offset = company_top_offset + 1.2 * cm_to_points # 在公司名称行下方 1.2cm
    
    # 添加文本框 (使用矩形 msoShapeRectangle = 1)
    date_textbox = doc.Shapes.AddShape(1, inner_left  + 8.5 * cm_to_points, inner_top + date_top_offset, sign_width, sign_height)
    
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
    text_frame_2.TextRange.ParagraphFormat.TabStops.Add(Position=width - 0.25 * cm_to_points, Alignment=2, Leader=0)

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

    # 在文字下方创建一个新的矩形框（主体内容区域）
    # 预留标题高度约 1.15cm
    header_height = 1.15 * cm_to_points
    # 内部矩形与外部矩形的间距：0.5cm
    padding = 0.5 * cm_to_points

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


    # 保存文档
    # FileFormat=12 代表 docx 格式 (wdFormatXMLDocument)
    doc.SaveAs(str(file_path), FileFormat=12)
    # 关闭文档和 Word 应用程序
    doc.Close()
    word.Quit()


if __name__ == '__main__':
    create_document(file_path=Path(__file__).parent.parent / 'source' / 'test.docx', item={})
