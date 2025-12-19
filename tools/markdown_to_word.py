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

# 导入新配置系统
from config import ConfigManager, StyleConfig
from converters.markdown_parser import MarkdownParser
from converters.word_generator import WordGenerator


class SmartDocGeneratorTool(Tool):
    """智能文档生成工具（重构版）"""
    
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        try:
            # 1. 提取参数
            markdown_text = tool_parameters.get('markdown_text', '')
            if not markdown_text:
                yield self.create_json_message({"error": "Markdown文本不能为空"})
                return
            
            theme = tool_parameters.get('templates', 'default')
            style_config_json = tool_parameters.get('style_config', '')
            enable_charts = tool_parameters.get('enable_charts', False)
            chart_data = tool_parameters.get('chart_data', '')
            
            # 2. 加载配置（统一入口）
            config_manager = ConfigManager()
            config = config_manager.load_config(
                json_config=style_config_json,
                theme=theme
            )
            
            # 3. 提取标题作为文件名
            output_file = self._extract_filename(markdown_text)
            
            # 4. 创建临时文件
            with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_file:
                temp_output_path = temp_file.name
            
            try:
                # 5. 初始化生成器
                markdown_parser = MarkdownParser()
                word_generator = WordGenerator(
                    config=config,
                    enable_charts=enable_charts,
                    chart_data=chart_data
                )
                
                # 6. 解析和生成
                parsed_content = markdown_parser.parse(markdown_text)
                success = word_generator.generate(
                    parsed_content,
                    temp_output_path,
                    markdown_text=markdown_text
                )
                
                if success:
                    # 读取生成的文件
                    with open(temp_output_path, 'rb') as f:
                        file_content = f.read()
                    
                    # 返回文件
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
                        "theme": theme,
                        "charts_enabled": enable_charts
                    })
                else:
                    yield self.create_json_message({"error": "Word文档生成失败"})
                    
            finally:
                # 清理临时文件
                if os.path.exists(temp_output_path):
                    os.unlink(temp_output_path)
                    
        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            yield self.create_json_message({
                "error": f"转换过程中发生错误: {str(e)}",
                "detail": error_detail
            })
    
    def _extract_filename(self, markdown_text: str) -> str:
        """从 markdown 文本中提取第一个一级标题作为文件名
        
        Args:
            markdown_text: Markdown 文本
        
        Returns:
            文件名（包含 .docx 扩展名）
        """
        # 处理转义的换行符
        text = markdown_text.replace('\\n', '\n')
        
        # 按行分割，只检查前10行
        lines = text.split('\n')
        for line in lines[:10]:
            line = line.strip()
            # 严格匹配：行首必须是 # 后跟空格，且不是 ##
            if line.startswith('# ') and not line.startswith('##'):
                title = line[2:].strip()
                # 确保标题不为空，且长度合理
                if title and '\n' not in title and '\r' not in title and len(title) <= 200:
                    # 如果标题中包含 ##，截取到 ## 之前
                    if '##' in title:
                        title = title.split('##')[0].strip()
                    
                    # 清理文件名
                    return self._sanitize_filename(title) + '.docx'
        
        # 如果没有找到标题，使用默认文件名
        return 'output.docx'
    
    def _sanitize_filename(self, filename: str) -> str:
        """清理文件名，移除不允许的字符
        
        Args:
            filename: 原始文件名
        
        Returns:
            清理后的文件名
        """
        # 移除不允许的文件名字符: < > : " / \ | ? *
        invalid_chars = r'[<>:"/\\|?*]'
        filename = re.sub(invalid_chars, '_', filename)
        
        # 移除换行符和制表符
        filename = filename.replace('\n', '_').replace('\r', '_').replace('\t', '_')
        
        # 移除前后空格和点
        filename = filename.strip(' .')
        
        # 移除连续的下划线
        filename = re.sub(r'_+', '_', filename)
        
        # 限制长度
        if len(filename) > 200:
            filename = filename[:200]
        
        return filename if filename else 'output'
