# 常见问题与解决

## 安装问题

### 问题1：pip install python-pptx 失败

**症状**：安装时报错

**解决**：
```bash
# 升级pip
python -m pip install --upgrade pip

# 然后安装
pip install python-pptx
```

---

### 问题2：导入时报错 `No module named 'pptx'`

**症状**：
```
ModuleNotFoundError: No module named 'pptx'
```

**解决**：
```bash
# 确认已安装
pip show python-pptx

# 如果没找到，重新安装
pip install python-pptx

# 如果使用的是虚拟环境，确保在正确的环境中安装
```

---

## 文件操作问题

### 问题3：文件保存失败 `PermissionError`

**症状**：
```
PermissionError: [Errno 13] Permission denied: 'output.pptx'
```

**解决**：
1. 关闭正在打开的PPT文件
2. 检查文件是否被其他程序占用
3. 使用不同的文件名
4. 检查文件夹权限

---

### 问题4：找不到文件 `FileNotFoundError`

**症状**：
```
FileNotFoundError: [Errno 2] No such file or directory: 'image.png'
```

**解决**：
```python
# 使用绝对路径
import os
abs_path = os.path.abspath('image.png')

# 或检查文件是否存在
if os.path.exists('image.png'):
    print("文件存在")
else:
    print("文件不存在")
```

---

## 格式问题

### 问题5：图片显示不出来

**可能原因**：
1. 图片路径错误
2. 图片格式不支持
3. 图片太大

**解决**：
```python
# 使用支持的格式：PNG, JPEG, GIF, BMP
pic = slide.shapes.add_picture(
    'image.png',  # 建议用PNG
    left=Inches(1),
    top=Inches(2),
    width=Inches(4)
)

# 检查图片是否有效
from PIL import Image
img = Image.open('image.png')
print(f"图片格式: {img.format}, 尺寸: {img.size}")
```

---

### 问题6：中文显示乱码

**可能原因**：字体不支持中文

**解决**：
```python
# 指定支持中文的字体
from pptx.util import Pt

p.font.name = 'Microsoft YaHei'  # 或 'SimHei', 'Arial Unicode MS'
p.font.size = Pt(14)
```

---

### 问题7：图表显示异常

**可能原因**：
1. 数据格式不对
2. 图表类型不支持

**解决**：
```python
# 确保数据格式正确
chart_data = CategoryChartData()
chart_data.categories = ['Q1', 'Q2', 'Q3']  # 必须是列表
chart_data.add_series('销售', (10, 20, 30))  # 必须是元组或列表

# 检查图表类型
from pptx.enum.chart import XL_CHART_TYPE
print("支持的图表类型:", [t for t in dir(XL_CHART_TYPE) if not t.startswith('_')])
```

---

## 导出问题

### 问题8：导出PDF失败

**症状**：
```
com_error: (-2147352567, '异常', (0, 'Microsoft PowerPoint', '文件未保存', 0, -2147024891))
```

**可能原因**：
1. 未安装Microsoft PowerPoint
2. 文件路径有中文

**解决**：
```python
# 方法1：安装Microsoft PowerPoint

# 方法2：使用英文路径
prs.save('C:/temp/output.pdf')

# 方法3：使用comtypes导出图片再转PDF
```

---

### 问题9：导出图片失败

**解决**：
```python
# 需要pywin32（仅Windows）
pip install pywin32

# 或者使用其他方法导出
```

---

## 性能问题

### 问题10：大文件处理很慢

**解决**：
```python
# 避免频繁保存
prs = Presentation()
# ... 添加很多内容 ...
prs.save('output.pptx')  # 最后一次性保存

# 不要在循环中保存
```

---

### 问题11：内存占用太高

**解决**：
```python
# 处理完后及时清理
del prs
import gc
gc.collect()
```

---

## 其他问题

### 问题12：如何查看PPT结构？

```python
from scripts.export import get_slide_info

info = get_slide_info('file.pptx')
print(f"共 {info['slide_count']} 页")
for slide in info['slides']:
    print(f"  第{slide['index']}页: {slide.get('title', '无标题')}")
```

---

### 问题13：如何获取所有可用布局？

```python
from pptx import Presentation

prs = Presentation()
for i, layout in enumerate(prs.slide_layouts):
    print(f"{i}: {layout.name}")
```

---

### 问题14：如何复制幻灯片？

```python
# 复制幻灯片内容
source_slide = prs.slides[0]
new_slide = prs.slides.add_slide(prs.slide_layouts[1])

# 复制形状
for shape in source_slide.shapes:
    # 注意：无法直接复制，需要手动创建
    pass
```

---

## 获取帮助

- 官方文档：https://python-pptx.readthedocs.io/
- GitHub：https://github.com/python-openxml/python-pptx
- 问题反馈：在skill页面评论