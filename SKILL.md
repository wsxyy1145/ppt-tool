---
name: ppt-tool
description: |
  多功能PPT操作工具，支持创建、编辑、格式调整、导出等功能。使用python-pptx库进行PPT操作。
  适用于以下场景：
  (1) 创建新的PPT演示文稿
  (2) 打开并编辑现有的PPT文件
  (3) 添加幻灯片、文字、图片、图表等内容
  (4) 调整幻灯片布局、样式和主题
  (5) 导出为PDF、图片或视频格式
  (6) 添加简单的动画效果
  触发词：PPT、幻灯片、演示文稿、做PPT、PowerPoint、presentation
---

# PPT Tool - 多功能PPT操作技能

本技能使用python-pptx库进行PPT操作，提供创建、编辑、格式调整和导出等功能。

## 环境准备

### 1. 安装依赖

```bash
pip install python-pptx Pillow img2pdf reportlab
```

- `python-pptx`: PPT操作核心库
- `Pillow`: 图片处理（导出为图片时需要）
- `img2pdf`: 图片转PDF（导出为PDF时需要）
- `reportlab`: PDF生成（高级PDF导出）

### 2. 基本导入

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
```

## 核心操作

### 创建新PPT

```python
from pptx import Presentation

# 创建空白PPT（默认16:9比例）
prs = Presentation()

# 设置幻灯片尺寸（默认宽屏16:9）
prs.slide_width = Inches(10)
prs.slide_height = Inches(5.625)

# 添加标题页
title_slide_layout = prs.slide_layouts[0]  # 标题页布局
slide = prs.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]
title.text = "演示标题"
subtitle.text = "副标题内容"

# 保存文件
prs.save('output.pptx')
```

### 打开现有PPT

```python
from pptx import Presentation

# 打开现有文件
prs = Presentation('existing.pptx')

# 遍历所有幻灯片
for slide in prs.slides:
    print(f"Slide: {slide.slide_number}")

# 保存为新文件
prs.save('new_file.pptx')
```

### 添加幻灯片

```python
# 使用不同的布局添加幻灯片
# 布局索引：0=标题页, 1=标题+内容, 2=标题+内容（两列）, 3=仅标题, 4=空白, 5=内容+标题

# 标题+内容布局
slide = prs.slides.add_slide(prs.slide_layouts[1])
title = slide.shapes.title
title.text = "幻灯片标题"

# 在内容区域添加文本框
body_shape = slide.placeholders[1]
tf = body_shape.text_frame
tf.text = "第一段内容"

# 添加更多段落
p = tf.add_paragraph()
p.text = "第二段内容"
p.level = 1  # 缩进级别
```

### 添加文本

```python
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# 在指定位置添加文本框
left = Inches(1)
top = Inches(2)
width = Inches(8)
height = Inches(1)

textbox = slide.shapes.add_textbox(left, top, width, height)
tf = textbox.text_frame
p = tf.paragraphs[0]
p.text = "文本内容"
p.font.size = Pt(24)  # 字体大小
p.font.bold = True    # 粗体
p.font.color.rgb = RGBColor(255, 0, 0)  # 红色
p.alignment = PP_ALIGN.CENTER  # 居中对齐
```

### 添加图片

```python
from pptx.util import Inches

# 添加图片（保持比例）
left = Inches(1)
top = Inches(2)
width = Inches(4)

pic = slide.shapes.add_picture('image.png', left, top, width=width)

# 添加图片（指定宽高，会变形）
pic = slide.shapes.add_picture('image.png', left, top, width=Inches(4), height=Inches(3))
```

### 添加图表

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

# 创建图表数据
chart_data = CategoryChartData()
chart_data.categories = ['第一季度', '第二季度', '第三季度', '第四季度']
chart_data.add_series('销售额', (100, 120, 140, 180))

# 添加图表
x, y, cx, cy = Inches(1), Inches(2), Inches(6), Inches(4.5)
chart = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    x, y, cx, cy,
    chart_data
).chart
```

支持的图表类型：
- `XL_CHART_TYPE.COLUMN_CLUSTERED` - 柱状图
- `XL_CHART_TYPE.LINE` - 折线图
- `XL_CHART_TYPE.PIE` - 饼图
- `XL_CHART_TYPE.BAR_CLUSTERED` - 条形图

### 添加形状

```python
from pptx.enum.shapes import MSO_SHAPE

# 添加矩形
shape = slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    left=Inches(1), top=Inches(1),
    width=Inches(2), height=Inches(1)
)

# 其他形状：MSO_SHAPE.OVAL（椭圆）、MSO_SHAPE.TRIANGLE（三角形）等
shape.text = "形状内的文字"
```

## 格式调整

### 设置背景颜色

```python
from pptx.util import Inches
from pptx.dml.color import RGBColor

# 设置单色背景
background = slide.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = RGBColor(255, 255, 255)  # 白色
```

### 设置主题颜色

```python
# 设置占位符文字格式
placeholder = slide.placeholders[1]
tf = placeholder.text_frame
tf.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)  # 黑色文字
```

### 调整布局

```python
# 设置页边距
textbox = slide.shapes[0]
tf = textbox.text_frame
tf.margin_left = Inches(0.5)
tf.margin_right = Inches(0.5)
tf.margin_top = Inches(0.3)
tf.margin_bottom = Inches(0.3)

# 设置行距
p = tf.paragraphs[0]
p.space_before = Pt(12)
p.space_after = Pt(12)
```

### 设置表格

```python
from pptx.util import Inches

# 添加表格
rows = 3
cols = 4
left = Inches(1)
top = Inches(2)
width = Inches(8)
height = Inches(2)

table = slide.shapes.add_table(rows, cols, left, top, width, height).table

# 设置单元格内容
table.cell(0, 0).text = "标题1"
table.cell(0, 1).text = "标题2"

# 设置列宽
table.columns[0].width = Inches(2)
table.columns[1].width = Inches(2)
```

## 导出功能

### 导出为PDF

```python
# 方法1：直接保存为PDF（需要安装了Microsoft PowerPoint）
prs.save('output.pdf')  # 保存时选择PDF格式

# 方法2：使用图片+PDF转换
from PIL import Image
import img2pdf

# 将每张幻灯片导出为图片
output_folder = 'slides'
os.makedirs(output_folder, exist_ok=True)

for i, slide in enumerate(prs.slides):
    # 需要使用额外工具将slide导出为图片
    slide.save(os.path.join(output_folder, f'slide_{i}.png'))
```

### 导出为图片

```python
# 使用pywin32（仅Windows，需要安装PowerPoint）
import win32com.client
import os

ppt = win32com.client.Dispatch("PowerPoint.Application")
ppt.Visible = 1

prs = ppt.Presentations.Open(os.path.abspath('input.pptx'))
prs.SaveAs(os.path.abspath('output'), 2)  # 2 = ppSaveAsPNG
prs.Close()
ppt.Quit()
```

## 动画效果

python-pptx对动画的支持有限，但可以设置基本效果：

```python
# 注意：python-pptx对动画的支持非常有限
# 建议使用VBA或直接操作PPTX XML来添加复杂动画

# 简单的切换效果
slide.transition = 'fade'  # 需要手动设置XML
```

对于复杂动画，建议：
1. 先在PowerPoint中创建模板，设置好动画
2. 使用python-pptx修改内容，保持动画不变

## 常用布局索引

| 索引 | 布局名称 | 说明 |
|------|---------|------|
| 0 | Title Slide | 标题页 |
| 1 | Title and Content | 标题+内容 |
| 2 | Section Header | 章节标题 |
| 3 | Two Content | 两列内容 |
| 4 | Comparison | 对比布局 |
| 5 | Title Only | 仅标题 |
| 6 | Blank | 空白 |
| 7 | Content with Caption | 内容+标题 |
| 8 | Picture with Caption | 图片+标题 |

## 完整示例：创建产品介绍PPT

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os

def create_product_presentation():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    # 封面页
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = "产品介绍"
    subtitle.text = "2024年度新品发布"

    # 产品概述页
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "产品概述"
    body = slide.placeholders[1].text_frame
    body.text = "产品亮点："
    p = body.add_paragraph()
    p.text = "• 高性能处理器"
    p.level = 1
    p = body.add_paragraph()
    p.text = "• 创新设计"
    p.level = 1
    p = body.add_paragraph()
    p.text = "• 优质用户体验"
    p.level = 1

    # 特性详情页
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "核心特性"
    body = slide.placeholders[1].text_frame
    body.text = "1. 强大性能"
    p = body.add_paragraph()
    p.text = "采用最新一代处理器，运行速度提升50%"
    p.level = 1
    p = body.add_paragraph()
    p.text = "2. 精美设计"
    p.level = 1

    # 团队页
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "团队介绍"
    body = slide.placeholders[1].text_frame
    body.text = "核心团队成员："
    for name in ["张三 - CEO", "李四 - CTO", "王五 - 设计总监"]:
        p = body.add_paragraph()
        p.text = name
        p.level = 1

    # 结束页
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = "谢谢！"
    subtitle.text = "联系我们：info@example.com"

    # 保存
    output_path = 'product_presentation.pptx'
    prs.save(output_path)
    print(f"PPT已创建: {os.path.abspath(output_path)}")
    return output_path

if __name__ == '__main__':
    create_product_presentation()
```

## 注意事项

1. **文件路径**：使用绝对路径或确保相对路径正确
2. **字体兼容**：某些字体在不同系统可能显示不同
3. **图片格式**：建议使用PNG或JPEG格式
4. **PowerPoint依赖**：部分高级功能需要安装Microsoft PowerPoint

## 相关参考

- python-pptx官方文档：https://python-pptx.readthedocs.io/
- 图表类型参考：https://python-pptx.readthedocs.io/en/latest/user/charts.html
- 详细API：http://officeopenxml.com/anatomyofOOXML-ppt.php

---

## 参考文档

更多详细文档位于 `references/` 目录：

| 文件 | 说明 |
|------|------|
| `references/quick_start.md` | 5分钟快速入门 |
| `references/api_reference.md` | 完整API参考 |
| `references/examples.md` | 50+代码示例 |
| `references/troubleshooting.md` | 常见问题与解决 |

使用示例：

```python
# 快速创建PPT
from scripts.quick_create import create_simple_ppt

create_simple_ppt(
    title="演示文稿",
    subtitle="副标题",
    slides_content=["内容1", {"title": "页面2", "content": "内容2"}]
)

# 添加图片/图表
from scripts.add_content import add_image_to_slide, add_chart_to_slide

add_image_to_slide("photo.jpg", slide_index=0)
add_chart_to_slide(chart_data={"categories": ["A", "B"], "series": {"数据": (1, 2)}}, title="图表")

# 导出PDF
from scripts.export import ppt_to_pdf, get_slide_info

ppt_to_pdf("input.pptx")
print(get_slide_info("input.pptx"))
```