# -*- coding: utf-8 -*-
"""
测试Markdown解析器的基本功能
"""

import unittest
import sys
from pathlib import Path

# 添加src目录到Python路径
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root / "src"))

from converters.markdown_parser import MarkdownParser, MarkdownElement


class TestMarkdownParser(unittest.TestCase):
    """Markdown解析器测试类"""
    
    def setUp(self):
        """测试前的准备工作"""
        self.parser = MarkdownParser()
    
    def test_parse_simple_text(self):
        """测试解析简单文本"""
        markdown_text = "这是一个简单的段落。"
        result = self.parser.parse(markdown_text)
        
        self.assertIsNotNone(result)
        self.assertIsInstance(result, MarkdownElement)
        self.assertEqual(result.element_type, 'document')
        self.assertIn('这是一个简单的段落', result.content)
    
    def test_parse_headers(self):
        """测试解析标题"""
        markdown_text = "# 一级标题\n## 二级标题\n### 三级标题"
        result = self.parser.parse(markdown_text)
        
        self.assertIsNotNone(result)
        self.assertIsInstance(result, MarkdownElement)
        
        # 检查元数据中的标题信息
        metadata = self.parser.extract_metadata(markdown_text)
        self.assertIn('headings', metadata)
        self.assertEqual(len(metadata['headings']), 3)
    
    def test_parse_bold_italic(self):
        """测试解析粗体和斜体"""
        markdown_text = "这是**粗体文本**和*斜体文本*。"
        result = self.parser.parse(markdown_text)
        
        self.assertIsNotNone(result)
        self.assertIsInstance(result, MarkdownElement)
    
    def test_parse_lists(self):
        """测试解析列表"""
        markdown_text = """
- 第一项
- 第二项
- 第三项

1. 有序第一项
2. 有序第二项
"""
        result = self.parser.parse(markdown_text)
        
        self.assertIsNotNone(result)
        self.assertIsInstance(result, MarkdownElement)
    
    def test_parse_code_blocks(self):
        """测试解析代码块"""
        markdown_text = """
```python
def hello():
    print("Hello, World!")
```
"""
        result = self.parser.parse(markdown_text)
        
        self.assertIsNotNone(result)
        self.assertIsInstance(result, MarkdownElement)
        
        # 检查元数据中的代码块信息
        metadata = self.parser.extract_metadata(markdown_text)
        self.assertIn('code_blocks', metadata)
        self.assertEqual(len(metadata['code_blocks']), 1)
        self.assertEqual(metadata['code_blocks'][0]['language'], 'python')
    
    def test_parse_table(self):
        """测试解析表格"""
        markdown_text = """
| 姓名 | 年龄 | 职业 |
|------|------|------|
| 张三 | 25   | 工程师 |
| 李四 | 30   | 设计师 |
"""
        result = self.parser.parse(markdown_text)
        
        self.assertIsNotNone(result)
        self.assertIsInstance(result, MarkdownElement)
        
        # 检查元数据中的表格信息
        metadata = self.parser.extract_metadata(markdown_text)
        self.assertIn('tables', metadata)
        self.assertEqual(len(metadata['tables']), 1)
    
    def test_parse_links(self):
        """测试解析链接"""
        markdown_text = "这是一个[链接示例](https://www.example.com)。"
        result = self.parser.parse(markdown_text)
        
        self.assertIsNotNone(result)
        self.assertIsInstance(result, MarkdownElement)
        
        # 检查元数据中的链接信息
        metadata = self.parser.extract_metadata(markdown_text)
        self.assertIn('links', metadata)
        self.assertEqual(len(metadata['links']), 1)
        self.assertEqual(metadata['links'][0]['url'], 'https://www.example.com')
    
    def test_parse_images(self):
        """测试解析图片"""
        markdown_text = "![示例图片](https://example.com/image.png)"
        result = self.parser.parse(markdown_text)
        
        self.assertIsNotNone(result)
        self.assertIsInstance(result, MarkdownElement)
        
        # 检查元数据中的图片信息
        metadata = self.parser.extract_metadata(markdown_text)
        self.assertIn('images', metadata)
        self.assertEqual(len(metadata['images']), 1)
        self.assertEqual(metadata['images'][0]['src'], 'https://example.com/image.png')
    
    def test_parse_quote(self):
        """测试解析引用"""
        markdown_text = """
> 这是一个引用块。
> 可以包含多行内容。
"""
        result = self.parser.parse(markdown_text)
        
        self.assertIsNotNone(result)
        self.assertIsInstance(result, MarkdownElement)
    
    def test_parse_empty_content(self):
        """测试解析空内容"""
        markdown_text = ""
        result = self.parser.parse(markdown_text)
        
        self.assertIsNotNone(result)
        self.assertIsInstance(result, MarkdownElement)
    
    def test_parse_complex_document(self):
        """测试解析复杂文档"""
        markdown_text = """
# 主标题

这是一个包含多种元素的复杂文档。

## 子标题

这里有**粗体**和*斜体*文本。

### 列表示例

- 无序列表项1
- 无序列表项2
  - 嵌套项

1. 有序列表项1
2. 有序列表项2

### 代码示例

```python
def test_function():
    return "Hello, World!"
```

### 表格示例

| 列1 | 列2 |
|-----|-----|
| 值1 | 值2 |

### 链接和图片

[链接](https://example.com)
![图片](https://example.com/image.png)

> 这是一个引用块

---

最后一段文字。
"""
        result = self.parser.parse(markdown_text)
        
        self.assertIsNotNone(result)
        self.assertIsInstance(result, MarkdownElement)
        
        # 检查元数据
        metadata = self.parser.extract_metadata(markdown_text)
        self.assertGreater(len(metadata['headings']), 0)
        self.assertGreater(len(metadata['code_blocks']), 0)
        self.assertGreater(len(metadata['tables']), 0)
        self.assertGreater(len(metadata['links']), 0)
        self.assertGreater(len(metadata['images']), 0)


if __name__ == '__main__':
    # 运行测试
    unittest.main(verbosity=2)