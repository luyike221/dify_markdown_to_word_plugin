import os
import re
import tempfile
from collections.abc import Generator
from typing import Any
from pathlib import Path

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage


# 动态添加src路径到模块搜索路径
current_file = Path(__file__)
src_path = current_file.parent.parent / 'src'

if str(src_path) not in os.sys.path:
    os.sys.path.insert(0, str(src_path))

from converters.markdown_parser import MarkdownParser
from converters.word_generator import WordGenerator
from converters.style_handler import StyleHandler
from utils.config import ConfigManager, StyleConfig, FontConfig, PageConfig
from utils.file_handler import FileHandler


class MarkdownToWordTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        try:
            # 提取参数
            markdown_text = tool_parameters.get('markdown_text', '')
            templates = tool_parameters.get('templates', 'default')
            font_family = tool_parameters.get('font_family', '微软雅黑')
            font_size = tool_parameters.get('font_size', 12)
            line_spacing = tool_parameters.get('line_spacing', 1.5)
            page_margins = tool_parameters.get('page_margins', 2.5)
            paper_size = tool_parameters.get('paper_size', 'A4')
            add_page_numbers = tool_parameters.get('add_page_numbers', True)
            

            if not markdown_text:
                yield self.create_json_message({
                    "error": "Markdown文本不能为空"
                })
                return
            
            # 检查是否成功导入业务逻辑模块
            if not all([MarkdownParser, WordGenerator, StyleHandler, ConfigManager]):
                yield self.create_json_message({
                    "error": "业务逻辑模块导入失败，请检查依赖"
                })
                return
            
            # 从 markdown 文本中提取第一个一级标题作为文件名
            def extract_h1_title(text: str) -> str | None:
                """提取第一个一级标题（# 开头的）"""
                # 先处理转义的换行符
                text = text.replace('\\n', '\n')
                # 按行分割，只检查前几行（避免处理整个大文件）
                lines = text.split('\n')
                for line in lines[:10]:  # 只检查前10行
                    line = line.strip()
                    # 严格匹配：行首必须是 # 后跟空格，且不是 ##
                    if line.startswith('# ') and not line.startswith('##'):
                        title = line[2:].strip()
                        # 确保标题不为空，且长度合理（不超过200字符）
                        # 标题应该在一行内，不应该包含换行符
                        if title and '\n' not in title and '\r' not in title and len(title) <= 200:
                            # 进一步验证：标题不应该包含 markdown 的特殊字符（如 ##）
                            # 如果标题看起来像是包含了后续内容，则截取到第一个可能的结束位置
                            if '##' in title:
                                # 如果标题中包含 ##，说明可能误包含了后续内容，截取到 ## 之前
                                title = title.split('##')[0].strip()
                            return title
                return None
            
            def sanitize_filename(filename: str) -> str:
                """清理文件名，移除不允许的字符"""
                # 移除或替换不允许的文件名字符
                # Windows/Linux 不允许的字符: < > : " / \ | ? *
                invalid_chars = r'[<>:"/\\|?*]'
                filename = re.sub(invalid_chars, '_', filename)
                # 移除换行符和制表符
                filename = filename.replace('\n', '_').replace('\r', '_').replace('\t', '_')
                # 移除前后空格和点
                filename = filename.strip(' .')
                # 移除连续的下划线
                filename = re.sub(r'_+', '_', filename)
                # 限制长度（保留扩展名空间）
                if len(filename) > 200:
                    filename = filename[:200]
                return filename if filename else 'output'
            
            # 提取一级标题作为文件名
            h1_title = extract_h1_title(markdown_text)
            if h1_title:
                # 清理标题并添加扩展名
                clean_title = sanitize_filename(h1_title)
                if clean_title and clean_title != 'output' and len(clean_title) > 0:
                    output_file = f"{clean_title}.docx"
                else:
                    # 如果清理后为空，使用默认文件名
                    output_file = 'output.docx'
            else:
                # 如果没有一级标题，使用默认文件名
                output_file = 'output.docx'
            
            # 创建临时文件来保存输出
            with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_file:
                temp_output_path = temp_file.name
            
            try:
                # 初始化配置管理器
                config_manager = ConfigManager()
                
                # 创建自定义样式配置
                font_config = FontConfig(
                    name=font_family,
                    size=font_size
                )
                
                page_config = PageConfig(
                    margin_top=page_margins,
                    margin_bottom=page_margins,
                    margin_left=page_margins,
                    margin_right=page_margins
                )
                
                style_config = StyleConfig(
                    theme=templates,
                    font=font_config,
                    page=page_config,
                    line_spacing=line_spacing,
                    enable_page_numbers=add_page_numbers
                )
                
                # 初始化转换器组件
                style_handler = StyleHandler()
                markdown_parser = MarkdownParser()
                
                # 设置主题
                if templates and templates != 'default':
                    if not style_handler.set_theme(templates):
                        # 如果主题不存在，使用默认主题
                        style_handler.set_theme('default')
                else:
                    style_handler.set_theme('default')
                
                # 将 StyleHandler 传递给 WordGenerator
                word_generator = WordGenerator(style_handler=style_handler)
                
                # 解析Markdown
                parsed_content = markdown_parser.parse(markdown_text)
                
                # 生成Word文档
                success = word_generator.generate(parsed_content, temp_output_path)
                
                if success:
                    # 读取生成的文件
                    with open(temp_output_path, 'rb') as f:
                        file_content = f.read()
                    
                    # 返回文件到 files 列表（移除不兼容的 save_as 参数）
                    yield self.create_blob_message(
                        blob=file_content,
                        meta={
                            "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            "filename": output_file
                        }
                    )
                    
                    # 返回成功结果
                    yield self.create_json_message({
                        "result": "Word文档生成成功",
                        "output_file": output_file,
                        "file_size": len(file_content),
                        "settings": {
                            "font_family": font_family,
                            "font_size": font_size,
                            "line_spacing": line_spacing,
                            "page_margins": page_margins,
                            "paper_size": paper_size,
                            "theme": templates,
                            "page_numbers": add_page_numbers
                        }
                    })
                else:
                    yield self.create_json_message({
                        "error": "Word文档生成失败"
                    })
                    
            finally:
                # 清理临时文件
                if os.path.exists(temp_output_path):
                    os.unlink(temp_output_path)
                    
        except Exception as e:
            yield self.create_json_message({
                "error": f"转换过程中发生错误: {str(e)}"
            })
