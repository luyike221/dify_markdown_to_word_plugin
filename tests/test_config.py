# -*- coding: utf-8 -*-
"""
测试配置文件
定义测试环境的配置和常量
"""

import os
import tempfile
from pathlib import Path

# 测试根目录
TEST_ROOT = Path(__file__).parent
PROJECT_ROOT = TEST_ROOT.parent
SRC_ROOT = PROJECT_ROOT / "src"

# 测试数据目录
FIXTURES_DIR = TEST_ROOT / "fixtures"
TEST_DATA_DIR = FIXTURES_DIR / "test_data"
TEST_IMAGES_DIR = FIXTURES_DIR / "images"
TEST_STYLES_DIR = FIXTURES_DIR / "styles"

# 临时文件目录
TEMP_DIR = Path(tempfile.gettempdir()) / "md_to_word_tests"
TEMP_DIR.mkdir(exist_ok=True)

# 测试文件路径
TEST_FILES = {
    'simple_md': FIXTURES_DIR / 'simple.md',
    'complex_md': FIXTURES_DIR / 'complex.md',
    'table_md': FIXTURES_DIR / 'table.md',
    'code_md': FIXTURES_DIR / 'code.md',
    'image_md': FIXTURES_DIR / 'image.md',
    'empty_md': FIXTURES_DIR / 'empty.md',
    'invalid_md': FIXTURES_DIR / 'invalid.md',
}

# 测试输出文件
TEST_OUTPUT_FILES = {
    'simple_docx': TEMP_DIR / 'simple_output.docx',
    'complex_docx': TEMP_DIR / 'complex_output.docx',
    'table_docx': TEMP_DIR / 'table_output.docx',
    'code_docx': TEMP_DIR / 'code_output.docx',
    'image_docx': TEMP_DIR / 'image_output.docx',
}

# 测试样式配置
TEST_STYLES = {
    'default': {
        'font_name': '微软雅黑',
        'font_size': 12,
        'line_spacing': 1.15,
    },
    'academic': {
        'font_name': 'Times New Roman',
        'font_size': 12,
        'line_spacing': 1.5,
    },
    'business': {
        'font_name': 'Calibri',
        'font_size': 11,
        'line_spacing': 1.15,
    }
}

# 测试图片URL
TEST_IMAGE_URLS = {
    'valid_png': 'https://via.placeholder.com/300x200.png',
    'valid_jpg': 'https://via.placeholder.com/400x300.jpg',
    'invalid_url': 'https://invalid-url.com/image.png',
    'large_image': 'https://via.placeholder.com/2000x2000.png',
}

# 测试配置
TEST_CONFIG = {
    'timeout': 30,  # 测试超时时间（秒）
    'max_file_size': 10 * 1024 * 1024,  # 最大文件大小（10MB）
    'supported_formats': ['.md', '.markdown', '.txt'],
    'output_format': '.docx',
    'encoding': 'utf-8',
}

# 性能测试配置
PERFORMANCE_CONFIG = {
    'small_doc_size': 1000,  # 小文档字符数
    'medium_doc_size': 10000,  # 中等文档字符数
    'large_doc_size': 100000,  # 大文档字符数
    'max_conversion_time': 60,  # 最大转换时间（秒）
    'memory_limit': 500 * 1024 * 1024,  # 内存限制（500MB）
}

# 错误测试配置
ERROR_TEST_CONFIG = {
    'invalid_markdown': [
        '# Heading\n\n[Invalid link](',
        '![Invalid image](',
        '```\nUnclosed code block',
        '| Invalid | Table\n| --- |',
    ],
    'network_errors': [
        'ConnectionError',
        'TimeoutError',
        'HTTPError',
    ],
    'file_errors': [
        'FileNotFoundError',
        'PermissionError',
        'IsADirectoryError',
    ]
}

# 断言辅助函数
def assert_file_exists(file_path):
    """断言文件存在"""
    assert Path(file_path).exists(), f"文件不存在: {file_path}"

def assert_file_not_empty(file_path):
    """断言文件不为空"""
    path = Path(file_path)
    assert path.exists(), f"文件不存在: {file_path}"
    assert path.stat().st_size > 0, f"文件为空: {file_path}"

def assert_valid_docx(file_path):
    """断言是有效的docx文件"""
    from docx import Document
    try:
        doc = Document(file_path)
        assert len(doc.paragraphs) >= 0, "文档没有段落"
    except Exception as e:
        assert False, f"无效的docx文件: {e}"

def cleanup_temp_files():
    """清理临时文件"""
    import shutil
    if TEMP_DIR.exists():
        shutil.rmtree(TEMP_DIR)
    TEMP_DIR.mkdir(exist_ok=True)

# 测试数据生成器
class TestDataGenerator:
    """测试数据生成器"""
    
    @staticmethod
    def generate_markdown(size='small'):
        """生成指定大小的Markdown内容"""
        sizes = {
            'small': 500,
            'medium': 5000,
            'large': 50000,
        }
        
        target_size = sizes.get(size, 500)
        content = []
        
        # 添加标题
        content.append("# 测试文档")
        content.append("")
        content.append("这是一个自动生成的测试文档。")
        content.append("")
        
        # 添加内容直到达到目标大小
        section_num = 1
        while len('\n'.join(content)) < target_size:
            content.extend([
                f"## 第{section_num}节",
                "",
                f"这是第{section_num}节的内容。" * 10,
                "",
                "### 子节",
                "",
                "- 列表项1",
                "- 列表项2",
                "- 列表项3",
                "",
                "```python",
                "def test_function():",
                "    return 'Hello, World!'",
                "```",
                "",
            ])
            section_num += 1
        
        return '\n'.join(content)
    
    @staticmethod
    def generate_table_markdown(rows=5, cols=3):
        """生成包含表格的Markdown"""
        content = ["# 表格测试文档", "", "以下是一个测试表格：", ""]
        
        # 表格头
        headers = [f"列{i+1}" for i in range(cols)]
        content.append("| " + " | ".join(headers) + " |")
        content.append("|" + "---|" * cols)
        
        # 表格数据
        for i in range(rows):
            row_data = [f"数据{i+1}-{j+1}" for j in range(cols)]
            content.append("| " + " | ".join(row_data) + " |")
        
        return '\n'.join(content)
    
    @staticmethod
    def generate_code_markdown():
        """生成包含代码块的Markdown"""
        return """
# 代码测试文档

## Python代码

```python
def hello_world():
    print("Hello, World!")
    return "Hello, World!"

class TestClass:
    def __init__(self):
        self.value = 42
    
    def get_value(self):
        return self.value
```

## JavaScript代码

```javascript
function helloWorld() {
    console.log("Hello, World!");
    return "Hello, World!";
}

class TestClass {
    constructor() {
        this.value = 42;
    }
    
    getValue() {
        return this.value;
    }
}
```

## SQL代码

```sql
SELECT 
    id,
    name,
    email
FROM users
WHERE 
    active = 1
    AND created_at > '2023-01-01'
ORDER BY name ASC;
```
"""

# 环境检查
def check_test_environment():
    """检查测试环境"""
    issues = []
    
    # 检查必要的目录
    required_dirs = [TEST_ROOT, FIXTURES_DIR, TEMP_DIR]
    for dir_path in required_dirs:
        if not dir_path.exists():
            issues.append(f"缺少目录: {dir_path}")
    
    # 检查Python包
    required_packages = ['docx', 'markdown', 'beautifulsoup4']
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            issues.append(f"缺少Python包: {package}")
    
    return issues

if __name__ == "__main__":
    # 运行环境检查
    issues = check_test_environment()
    if issues:
        print("测试环境问题:")
        for issue in issues:
            print(f"  - {issue}")
    else:
        print("测试环境检查通过")