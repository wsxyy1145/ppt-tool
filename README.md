# PPT Tool / PPT工具

> 多功能PPT操作技能 / Multi-functional PPT manipulation skill

一个使用python-pptx的多功能PPT操作工具，支持创建、编辑、格式调整、导出等功能。
A multi-functional PowerPoint manipulation tool using python-pptx, supporting creation, editing, formatting, and export.

## 功能特性 / Features

- ✅ 创建/打开PPT文件 / Create/Open PPT files
- ✅ 添加幻灯片、文字、图片、图表 / Add slides, text, images, charts
- ✅ 调整布局和样式 / Adjust layout and styles
- ✅ 导出为PDF格式 / Export to PDF
- ✅ 添加形状和表格 / Add shapes and tables
- ✅ 支持多种图表类型（柱状图、折线图、饼图等）/ Support multiple chart types (bar, line, pie, etc.)
- ✅ 自定义字体颜色和样式 / Customize font colors and styles
- ✅ 批量处理幻灯片 / Batch process slides

## 适用场景 / Use Cases

- 自动化报告生成 / Automated report generation
- 批量制作演示文稿 / Batch create presentations
- 数据可视化 / Data visualization
- 教育课件制作 / Educational courseware
- 商务演示自动化 / Business presentation automation

## 快速开始 / Quick Start

### 1. 安装依赖 / Install Dependencies

```bash
pip install python-pptx Pillow img2pdf
```

**依赖说明 / Dependencies:**
- `python-pptx` - PPT操作核心库 / Core PPT manipulation library
- `Pillow` - 图片处理 / Image processing
- `img2pdf` - 图片转PDF / Image to PDF conversion

**可选依赖 / Optional:**
```bash
pip install reportlab  # 高级PDF导出 / Advanced PDF export
pip install pywin32    # Windows下导出图片 / Export images on Windows
```

### 2. 创建第一个PPT / Create Your First PPT

```python
from pptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "欢迎使用PPT工具"
slide.placeholders[1].text = "Welcome to PPT Tool"
prs.save("my_first_ppt.pptx")
```

### 3. 使用脚本快速创建 / Use Scripts for Quick Creation

```python
from scripts.quick_create import create_simple_ppt

create_simple_ppt(
    title="演示标题 / Presentation Title",
    subtitle="副标题 / Subtitle",
    slides_content=[
        "第一页内容 / Page 1 Content",
        {"title": "第二页标题", "content": "第二页内容"}
    ]
)
```

### 4. 添加图片和图表 / Add Images and Charts

```python
from scripts.add_content import add_image_to_slide, add_chart_to_slide

# 添加图片 / Add image
add_image_to_slide("photo.jpg", slide_index=0, width=4)

# 添加图表 / Add chart
chart_data = {
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series": {"销售额/Sales": (100, 120, 140, 180)}
}
add_chart_to_slide(chart_data, chart_type="column", title="销售报告/Sales Report")
```

### 5. 导出PDF / Export to PDF

```python
from scripts.export import ppt_to_pdf

# 需要安装Microsoft PowerPoint / Requires Microsoft PowerPoint installed
ppt_to_pdf("input.pptx", "output.pdf")
```

## 项目结构 / Project Structure

```
ppt-tool/
├── SKILL.md                 # 技能主文档 / Skill main documentation
├── _meta.json               # 元信息 / Meta information
├── README.md                # 项目说明 / Project readme
├── .gitignore               # Git忽略文件 / Git ignore file
├── scripts/                 # 实用脚本 / Utility scripts
│   ├── quick_create.py      # 快速创建PPT / Quick create PPT
│   ├── add_content.py       # 添加图片/图表/形状 / Add images/charts/shapes
│   └── export.py            # 导出功能 / Export functionality
└── references/              # 详细文档 / Detailed documentation
    ├── quick_start.md       # 5分钟入门 / 5-minute getting started
    ├── api_reference.md     # API参考 / API reference
    ├── examples.md          # 代码示例 / Code examples
    └── troubleshooting.md   # 常见问题 / Troubleshooting
```

## 触发词 / Trigger Words

`PPT` `幻灯片` `演示文稿` `做PPT` `PowerPoint` `presentation`

`Slide` `slides` `presentation` `PowerPoint`

## 文档 / Documentation

- [快速开始 / Quick Start](references/quick_start.md)
- [API参考 / API Reference](references/api_reference.md)
- [代码示例 / Examples](references/examples.md)
- [常见问题 / Troubleshooting](references/troubleshooting.md)

## 环境要求 / Requirements

- Python 3.7+
- Microsoft PowerPoint (用于PDF导出 / For PDF export)

## 常见问题 / FAQ

**Q: 如何导出PDF？/ How to export PDF?**
A: 需要安装Microsoft PowerPoint，然后使用 `ppt_to_pdf()` 函数。/ Requires Microsoft PowerPoint installed, then use `ppt_to_pdf()` function.

**Q: 支持哪些图片格式？/ What image formats are supported?**
A: PNG、JPEG、GIF、BMP等常见格式。/ PNG, JPEG, GIF, BMP and other common formats.

**Q: 支持哪些图表类型？/ What chart types are supported?**
A: 柱状图、折线图、饼图、条形图、面积图等。/ Column, line, pie, bar, area, etc.

## 更新日志 / Changelog

### v1.0.0 (2026-03-28)
- 初始版本发布 / Initial release
- 支持创建、编辑、导出PPT / Support create, edit, export PPT
- 提供快速脚本工具 / Provide quick script tools
- 完整的文档和示例 / Complete documentation and examples

## 许可证 / License

MIT License

## 作者 / Author

wsxyy1145 🦐