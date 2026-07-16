# Markdown 转 DOCX/PDF 示例文档

本文展示中文技术文档与 English technical terms 的混排效果，并包含 **加粗文本** 和 *斜体文本*。

## 1. 列表与引用

- 第一项：支持无序列表
- 第二项：支持中英文内容

1. 第一条有序内容
2. 第二条有序内容

> 这是引用块。它会转换为带左缩进的 Word 段落。

## 2. 代码与图片

```python
def convert(source: str) -> str:
    return "DOCX"
```

![示例图片](../image.png)

## 3. 表格

| 指标 | 中文说明 | Value |
| --- | --- | --- |
| 标题 | 自动映射 | Heading |
| 表格 | 三线表 | Native table |
