#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简单的转换测试脚本
用于验证Markdown到Word的基本转换功能
"""

import sys
from pathlib import Path

# 添加src目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root / "src"))

from converters.markdown_parser import MarkdownParser
from converters.word_generator import WordGenerator

def test_basic_conversion():
    """测试基本转换功能"""
    print("开始测试Markdown到Word转换...")
    
    # 创建解析器和生成器实例
    parser = MarkdownParser()
    generator = WordGenerator()
    
    # 测试文本
    markdown_text = """
# 测试文档

这是一个**测试文档**，包含多种*Markdown*元素。

## 功能列表

- 标题解析
- 文本格式化
- 列表处理
- 代码块支持

### 代码示例

```python
def hello_world():
    print("Hello, World!")
```

### 表格示例

| 功能 | 状态 |
|------|------|
| 解析 | ✅ |
| 转换 | 🚧 |

> 这是一个引用块，用于展示重要信息。

[访问GitHub](https://github.com)

![示例图片](https://example.com/image.png)
    """
    
    try:
        # 解析Markdown
        result = parser.parse(markdown_text)
        print(f"✅ 解析成功！文档类型: {result.element_type}")
        
        # 提取元数据
        metadata = parser.extract_metadata(markdown_text)
        print(f"✅ 元数据提取成功！")
        print(f"   - 标题数量: {len(metadata['headings'])}")
        print(f"   - 代码块数量: {len(metadata['code_blocks'])}")
        print(f"   - 表格数量: {len(metadata['tables'])}")
        print(f"   - 链接数量: {len(metadata['links'])}")
        print(f"   - 图片数量: {len(metadata['images'])}")
        
        # 显示标题信息
        print("\n📋 标题结构:")
        for heading in metadata['headings']:
            indent = "  " * (heading['level'] - 1)
            print(f"{indent}- {heading['title']} (H{heading['level']})")
        
        # 生成Word文档
        print("\n📄 开始生成Word文档...")
        output_path = "test_output.docx"
        
        # 由于当前的WordGenerator需要更完整的实现，我们先创建一个简单的Word文档
        success = generator.generate(result, output_path)
        
        if success:
            print(f"✅ Word文档生成成功！文件保存为: {output_path}")
            print(f"📁 文件路径: {Path(output_path).absolute()}")
        else:
            print("❌ Word文档生成失败")
        
        print("\n🎉 测试完成！Markdown到Word转换流程验证完毕。")
        return success
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_basic_conversion()
    sys.exit(0 if success else 1)