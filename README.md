# PPT Tool - 多功能PPT操作技能

一个使用python-pptx的多功能PPT操作工具，支持创建、编辑、格式调整、导出等功能。

## 功能特性

- ✅ 创建/打开PPT文件
- ✅ 添加幻灯片、文字、图片、图表
- ✅ 调整布局和样式
- ✅ 导出为PDF格式
- ✅ 添加形状和表格

## 快速开始

### 1. 安装依赖

```bash
pip install python-pptx Pillow img2pdf
```

### 2. 创建第一个PPT

```python
from pptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "欢迎使用PPT工具"
prs.save("my_first_ppt.pptx")
```

### 3. 使用脚本快速创建

```python
from scripts.quick_create import create_simple_ppt

create_simple_ppt(
    title="演示标题",
    subtitle="副标题",
    slides_content=["第一页", "第二页"]
)
```

## 项目结构

```
ppt-tool/
├── SKILL.md                 # 技能主文档
├── _meta.json               # 元信息
├── scripts/                 # 实用脚本
│   ├── quick_create.py      # 快速创建PPT
│   ├── add_content.py       # 添加图片/图表
│   └── export.py            # 导出功能
└── references/              # 详细文档
    ├── quick_start.md       # 5分钟入门
    ├── api_reference.md     # API参考
    ├── examples.md          # 代码示例
    └── troubleshooting.md   # 常见问题
```

## 触发词

PPT、幻灯片、演示文稿、做PPT、PowerPoint、presentation

## 文档

- [快速开始](references/quick_start.md)
- [API参考](references/api_reference.md)
- [代码示例](references/examples.md)
- [常见问题](references/troubleshooting.md)

## 依赖

- python-pptx
- Pillow
- img2pdf

## 许可证

MIT License

## 作者

小龙虾 🦐