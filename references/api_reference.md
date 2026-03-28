# API 参考

## Presentation 类

```python
from pptx import Presentation
```

### 主要方法

| 方法 | 说明 |
|------|------|
| `Presentation()` | 创建新的PPT对象 |
| `Presentation('file.pptx')` | 打开现有PPT |
| `slides.add_slide(layout)` | 添加新幻灯片 |
| `slide_layouts[index]` | 获取指定布局 |
| `save('file.pptx')` | 保存文件 |

### 属性

| 属性 | 说明 |
|------|------|
| `slide_width` | 幻灯片宽度（默认10英寸） |
| `slide_height` | 幻灯片高度（默认5.625英寸，即16:9） |

---

## 文本操作

### 添加文本框

```python
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# 创建文本框
textbox = slide.shapes.add_textbox(
    left=Inches(1),      # 左边距
    top=Inches(2),       # 上边距
    width=Inches(8),     # 宽度
    height=Inches(1)     # 高度
)

# 获取文本框
tf = textbox.text_frame
p = tf.paragraphs[0]
p.text = "文字内容"

# 设置格式
p.font.size = Pt(24)           # 字体大小
p.font.bold = True             # 粗体
p.font.italic = True           # 斜体
p.font.color.rgb = RGBColor(255, 0, 0)  # 红色
p.alignment = PP_ALIGN.CENTER  # 居中
```

### 段落格式

```python
# 添加段落
p = tf.add_paragraph()
p.text = "第二段"
p.level = 0                    # 级别（0=正文，1=一级缩进）
p.space_before = Pt(12)        # 段前间距
p.space_after = Pt(12)         # 段后间距
```

---

## 图片操作

```python
from pptx.util import Inches

# 添加图片（保持比例）
pic = slide.shapes.add_picture(
    'image.png',
    left=Inches(1),
    top=Inches(2),
    width=Inches(4)      # 指定宽度，高度自动
)

# 添加图片（指定宽高）
pic = slide.shapes.add_picture(
    'image.png',
    left=Inches(1),
    top=Inches(2),
    width=Inches(4),
    height=Inches(3)
)
```

---

## 图表操作

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

# 创建数据
chart_data = CategoryChartData()
chart_data.categories = ['A', 'B', 'C', 'D']
chart_data.add_series('系列1', (10, 20, 30, 40))

# 添加图表
chart = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    x=Inches(1), y=Inches(2),
    cx=Inches(6), cy=Inches(4.5),
    chart_data=chart_data
).chart
```

### 图表类型

| 类型 | 值 |
|------|------|
| 柱状图 | `XL_CHART_TYPE.COLUMN_CLUSTERED` |
| 折线图 | `XL_CHART_TYPE.LINE` |
| 饼图 | `XL_CHART_TYPE.PIE` |
| 条形图 | `XL_CHART_TYPE.BAR_CLUSTERED` |
| 面积图 | `XL_CHART_TYPE.AREA` |

---

## 形状操作

```python
from pptx.enum.shapes import MSO_SHAPE

# 添加形状
shape = slide.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,    # 形状类型
    left=Inches(1),
    top=Inches(1),
    width=Inches(2),
    height=Inches(1)
)
shape.text = "形状内的文字"

# 常用形状
MSO_SHAPE.RECTANGLE    # 矩形
MSO_SHAPE.OVAL         # 椭圆
MSO_SHAPE.TRIANGLE     # 三角形
MSO_SHAPE.ARROW_RIGHT  # 右箭头
MSO_SHAPE.STAR_5       # 五角星
```

---

## 表格操作

```python
from pptx.util import Inches

# 添加表格
table = slide.shapes.add_table(
    rows=3,                # 行数
    cols=4,                # 列数
    left=Inches(1),
    top=Inches(2),
    width=Inches(8),
    height=Inches(2)
).table

# 操作单元格
table.cell(0, 0).text = "标题"
table.cell(0, 0].text_frame.paragraphs[0].font.bold = True

# 设置列宽
table.columns[0].width = Inches(2)
```

---

## 颜色参考

```python
from pptx.dml.color import RGBColor

# 常用颜色
RGBColor(0, 0, 0)         # 黑色
RGBColor(255, 255, 255)   # 白色
RGBColor(255, 0, 0)       # 红色
RGBColor(0, 255, 0)       # 绿色
RGBColor(0, 0, 255)       # 蓝色
RGBColor(255, 255, 0)     # 黄色
RGBColor(128, 128, 128)   # 灰色
```

---

## 布局索引

| 索引 | 布局 | 说明 |
|------|------|------|
| 0 | Title Slide | 标题页 |
| 1 | Title and Content | 标题+内容 |
| 2 | Section Header | 章节标题 |
| 3 | Two Content | 两列内容 |
| 4 | Comparison | 对比 |
| 5 | Title Only | 仅标题 |
| 6 | Blank | 空白 |
| 7 | Content with Caption | 内容+标题 |
| 8 | Picture with Caption | 图片+标题 |