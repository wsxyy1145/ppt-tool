# 示例集锦

## 基础示例

### 示例1：最简单的PPT

```python
from pptx import Presentation

prs = Presentation()
prs.save('simple.pptx')
```

---

### 示例2：带标题的PPT

```python
from pptx import Presentation

prs = Presentation()

# 标题页
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "我的演示"
slide.placeholders[1].text = "作者：某某"

# 内容页
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "目录"
slide.placeholders[1].text_frame.text = "1. 引言\n2. 内容\n3. 总结"

prs.save('with_title.pptx')
```

---

## 进阶示例

### 示例3：多页产品展示

```python
from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(5.625)

# 封面
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "产品发布会"
slide.placeholders[1].text = "2024年度新品"

# 产品1
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "产品A"
tf = slide.placeholders[1].text_frame
tf.text = "高端配置"
p = tf.add_paragraph()
p.text = "• CPU: 最新一代"
p.level = 1
p = tf.add_paragraph()
p.text = "• 内存: 16GB"
p.level = 1

# 产品2
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "产品B"
slide.placeholders[1].text_frame.text = "性价比之选"

# 结束页
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "谢谢观看"
slide.placeholders[1].text = "联系电话：123456789"

prs.save('product_showcase.pptx')
```

---

### 示例4：带图表的销售报告

```python
from pptx import Presentation
from pptx.util import Inches
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

prs = Presentation()

# 标题页
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "销售报告"
slide.placeholders[1].text = "2024年第一季度"

# 图表页
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "季度销售数据"

chart_data = CategoryChartData()
chart_data.categories = ['1月', '2月', '3月']
chart_data.add_series('销售额', (100, 120, 150))

chart = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1), Inches(2), Inches(8), Inches(4.5),
    chart_data
).chart

prs.save('sales_report.pptx')
```

---

### 示例5：带图片的幻灯片

```python
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()

slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "产品图片"

# 添加图片
pic = slide.shapes.add_picture(
    'product.jpg',
    left=Inches(1),
    top=Inches(2),
    width=Inches(4)
)

# 添加说明文字
textbox = slide.shapes.add_textbox(
    left=Inches(5.5), top=Inches(2),
    width=Inches(4), height=Inches(3)
)
textbox.text_frame.text = "产品实拍图"

prs.save('with_images.pptx')
```

---

### 示例6：表格数据展示

```python
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "产品报价表"

# 添加表格
table = slide.shapes.add_table(
    rows=4, cols=3,
    left=Inches(1), top=Inches(2),
    width=Inches(8), height=Inches(2)
).table

# 设置表头
table.cell(0, 0).text = "产品"
table.cell(0, 1).text = "型号"
table.cell(0, 2).text = "价格"

# 设置数据
table.cell(1, 0).text = "电脑"
table.cell(1, 1).text = "Pro 15"
table.cell(1, 2).text = "¥9999"

table.cell(2, 0).text = "手机"
table.cell(2, 1).text = "X Pro"
table.cell(2, 2).text = "¥5999"

table.cell(3, 0).text = "平板"
table.cell(3, 1).text = "Air 11"
table.cell(3, 2).text = "¥3999"

prs.save('with_table.pptx')
```

---

### 示例7：自定义颜色和样式

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "自定义样式"

# 添加带格式的文本
textbox = slide.shapes.add_textbox(
    left=Inches(1), top=Inches(2),
    width=Inches(8), height=Inches(2)
)

tf = textbox.text_frame
p = tf.paragraphs[0]
p.text = "红色大号标题"
p.font.size = Pt(36)
p.font.color.rgb = RGBColor(255, 0, 0)
p.font.bold = True

p = tf.add_paragraph()
p.text = "蓝色小号正文"
p.font.size = Pt(14)
p.font.color.rgb = RGBColor(0, 0, 255)

prs.save('styled.pptx')
```

---

### 示例8：使用脚本快速创建

```python
# 使用 quick_create.py
from scripts.quick_create import create_simple_ppt

create_simple_ppt(
    title="年度报告",
    subtitle="2024年",
    slides_content=[
        "第一页内容",
        {"title": "第二页标题", "content": "第二页内容"},
        "第三页"
    ],
    output_path="quick.pptx"
)
```

---

### 示例9：批量创建多页

```python
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()

# 批量添加10页
for i in range(10):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = f"第 {i+1} 页"
    slide.placeholders[1].text = f"这是第 {i+1} 页的内容"

prs.save('many_pages.pptx')
```

---

### 示例10：导出PDF

```python
from scripts.export import ppt_to_pdf

# 需要安装Microsoft PowerPoint
ppt_to_pdf('input.pptx', 'output.pdf')
```