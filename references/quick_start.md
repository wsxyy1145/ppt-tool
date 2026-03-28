# PPT工具 - 快速开始指南

## 5分钟快速入门

### 1. 安装依赖

```bash
pip install python-pptx Pillow img2pdf
```

### 2. 创建你的第一个PPT

```python
from pptx import Presentation

# 创建PPT
prs = Presentation()

# 添加标题页
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "欢迎使用PPT工具"
slide.placeholders[1].text = "副标题"

# 添加内容页
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "第一页"
slide.placeholders[1].text_frame.text = "这是内容"

# 保存
prs.save("my_first_ppt.pptx")
```

### 3. 运行脚本

```bash
python scripts/quick_create.py
```

---

## 常用场景

### 场景1：创建产品介绍PPT

```python
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()

# 封面
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "产品名称"
slide.placeholders[1].text = "副标题"

# 特点
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "产品特点"
slide.placeholders[1].text_frame.text = "• 特点1\n• 特点2\n• 特点3"

# 价格
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "价格"
slide.placeholders[1].text_frame.text = "¥2999"

prs.save("product.pptx")
```

### 场景2：批量添加图片

```python
from scripts.add_content import add_image_to_slide

# 给第一页添加图片
add_image_to_slide(
    image_path="photo.jpg",
    slide_index=0,
    left=1, top=2, width=4,
    output_path="with_image.pptx"
)
```

### 场景3：添加图表

```python
from scripts.add_content import add_chart_to_slide

chart_data = {
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series": {"销售额": (100, 150, 200, 250)}
}

add_chart_to_slide(
    chart_data=chart_data,
    chart_type="column",
    title="年度销售",
    output_path="with_chart.pptx"
)
```

---

## 使用脚本模板

### quick_create.py - 快速创建

```python
from scripts.quick_create import create_simple_ppt

create_simple_ppt(
    title="演示标题",
    subtitle="副标题",
    slides_content=[
        "第一页内容",
        {"title": "第二页", "content": "内容"}
    ]
)
```

### add_content.py - 添加内容

```python
# 添加图片
add_image_to_slide("图片路径", slide_index=0)

# 添加图表
add_chart_to_slide(chart_data, chart_type="column")

# 添加形状
add_shape("rectangle", text="文字")
```

### export.py - 导出

```python
from scripts.export import ppt_to_pdf, get_slide_info

# 导出PDF（需要PowerPoint）
ppt_to_pdf("input.pptx", "output.pdf")

# 查看PPT信息
info = get_slide_info("input.pptx")
print(f"共 {info['slide_count']} 页")
```

---

## 下一步

- 查看 `references/api_reference.md` 了解完整API
- 查看 `references/examples.md` 更多示例
- 查看 `references/troubleshooting.md` 常见问题