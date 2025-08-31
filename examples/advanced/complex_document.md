# 复杂文档示例

这是一个复杂的Markdown文档示例，展示了高级格式化功能和复杂结构。

## 目录

- [1. 介绍](#1-介绍)
- [2. 技术规格](#2-技术规格)
- [3. 代码示例](#3-代码示例)
- [4. 数据分析](#4-数据分析)
- [5. 结论](#5-结论)

---

## 1. 介绍

本文档展示了Markdown到Word转换器的**高级功能**，包括：

- 复杂的表格结构
- 多种编程语言的代码高亮
- 数学公式（如果支持）
- 嵌套的引用和列表
- 任务列表
- 脚注引用

### 1.1 项目背景

> **重要提示**：这是一个演示文档，用于测试转换器的各种功能。
> 
> 在实际使用中，请根据具体需求调整样式和格式。

### 1.2 使用场景

该转换器适用于以下场景：

- [ ] 技术文档编写
- [x] 学术论文撰写
- [x] 商务报告制作
- [ ] 个人笔记整理

## 2. 技术规格

### 2.1 系统要求

| 组件 | 最低要求 | 推荐配置 | 备注 |
|------|----------|----------|------|
| Python | 3.7+ | 3.9+ | 必需 |
| 内存 | 2GB | 4GB+ | 处理大文档时 |
| 存储 | 100MB | 500MB+ | 包含依赖库 |
| 操作系统 | Windows 10<br>macOS 10.14<br>Ubuntu 18.04 | 最新版本 | 跨平台支持 |

### 2.2 依赖库

```bash
# 核心依赖
pip install python-docx>=0.8.11
pip install markdown>=3.4.1
pip install beautifulsoup4>=4.11.1

# 可选依赖
pip install pillow>=9.0.0  # 图片处理
pip install requests>=2.28.0  # 网络图片下载
pip install pyyaml>=6.0  # 配置文件解析
```

## 3. 代码示例

### 3.1 Python代码

```python
class MarkdownToWordConverter:
    """Markdown到Word转换器主类"""
    
    def __init__(self, style_config=None):
        self.style_config = style_config or self._load_default_style()
        self.parser = MarkdownParser()
        self.generator = WordGenerator(self.style_config)
    
    def convert_file(self, input_path, output_path):
        """转换单个文件"""
        try:
            with open(input_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 解析Markdown
            parsed_content = self.parser.parse(content)
            
            # 生成Word文档
            document = self.generator.generate(parsed_content)
            
            # 保存文档
            document.save(output_path)
            
            return True
        except Exception as e:
            print(f"转换失败: {e}")
            return False
    
    def _load_default_style(self):
        """加载默认样式配置"""
        return {
            'font_name': '微软雅黑',
            'font_size': 12,
            'line_spacing': 1.15
        }
```

### 3.2 JavaScript代码

```javascript
// 前端集成示例
class MarkdownConverter {
    constructor(apiEndpoint) {
        this.apiEndpoint = apiEndpoint;
    }
    
    async convertToWord(markdownContent, options = {}) {
        const response = await fetch(`${this.apiEndpoint}/convert`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                content: markdownContent,
                style: options.style || 'default',
                format: 'docx'
            })
        });
        
        if (!response.ok) {
            throw new Error(`转换失败: ${response.statusText}`);
        }
        
        return await response.blob();
    }
}

// 使用示例
const converter = new MarkdownConverter('/api/v1');
converter.convertToWord(markdownText)
    .then(blob => {
        // 下载文件
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'document.docx';
        a.click();
    })
    .catch(error => console.error('转换错误:', error));
```

### 3.3 配置文件示例

```yaml
# config.yaml
converter:
  default_style: "academic"
  output_format: "docx"
  
styles:
  academic:
    font_family: "Times New Roman"
    font_size: 12
    line_spacing: 1.5
    margins:
      top: 2.54
      bottom: 2.54
      left: 3.18
      right: 3.18
  
  business:
    font_family: "Calibri"
    font_size: 11
    line_spacing: 1.15
    
features:
  table_of_contents: true
  page_numbers: true
  syntax_highlighting: true
```

## 4. 数据分析

### 4.1 性能测试结果

以下是不同文档大小的转换性能测试：

| 文档大小 | 页数 | 转换时间 | 内存使用 | CPU使用率 |
|----------|------|----------|----------|----------|
| 小型 | 1-5页 | 0.5-1.2秒 | 50-80MB | 15-25% |
| 中型 | 6-20页 | 1.2-3.5秒 | 80-150MB | 25-40% |
| 大型 | 21-50页 | 3.5-8.0秒 | 150-300MB | 40-60% |
| 超大型 | 50+页 | 8.0-20秒 | 300-500MB | 60-80% |

### 4.2 功能支持矩阵

| 功能 | 基础版 | 标准版 | 专业版 |
|------|--------|--------|--------|
| 基本格式 | ✅ | ✅ | ✅ |
| 表格 | ✅ | ✅ | ✅ |
| 图片 | ✅ | ✅ | ✅ |
| 代码高亮 | ❌ | ✅ | ✅ |
| 数学公式 | ❌ | ❌ | ✅ |
| 自定义样式 | ❌ | ✅ | ✅ |
| 批量转换 | ❌ | ✅ | ✅ |
| API接口 | ❌ | ❌ | ✅ |

### 4.3 错误统计

> **注意**：以下是常见的转换错误及其解决方案。

#### 4.3.1 图片相关错误

```python
# 错误处理示例
def handle_image_error(image_url, error):
    """处理图片加载错误"""
    error_types = {
        'NetworkError': '网络连接失败，请检查网络设置',
        'FileNotFound': '图片文件不存在，请检查路径',
        'FormatError': '不支持的图片格式，请使用PNG、JPG或GIF',
        'SizeError': '图片文件过大，请压缩后重试'
    }
    
    return error_types.get(type(error).__name__, '未知错误')
```

#### 4.3.2 样式冲突

当多个样式规则冲突时，系统按以下优先级处理：

1. **用户自定义样式** - 最高优先级
2. **主题样式** - 中等优先级  
3. **默认样式** - 最低优先级

## 5. 结论

### 5.1 主要优势

- ✨ **高保真转换**：保持原始格式和样式
- 🚀 **高性能**：快速处理大型文档
- 🎨 **灵活样式**：支持多种预设主题
- 🔧 **易于集成**：提供简单的API接口

### 5.2 使用建议

> 💡 **最佳实践**
> 
> 1. 使用标准的Markdown语法以确保最佳兼容性
> 2. 对于复杂表格，建议使用HTML表格语法
> 3. 图片建议使用相对路径或CDN链接
> 4. 定期更新依赖库以获得最新功能

### 5.3 未来规划

- [ ] 支持更多输出格式（PDF、HTML等）
- [ ] 增加实时预览功能
- [ ] 提供Web界面
- [ ] 支持协作编辑
- [ ] 集成云存储服务

---

**文档版本**：v1.0.0  
**最后更新**：2024年1月  
**作者**：Markdown转换器开发团队  
**联系方式**：support@example.com

---

*本文档使用Markdown编写，通过转换器生成Word格式。*