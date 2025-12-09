#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word文档生成器模块
负责将解析后的Markdown内容转换为Word文档
"""

import os
import re
from typing import Dict, List, Any, Optional, Union
from pathlib import Path
from dataclasses import dataclass

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.shared import OxmlElement, qn
    from docx.oxml.ns import nsdecls
    from docx.oxml import parse_xml
except ImportError:
    raise ImportError("请安装python-docx库: pip install python-docx")

from .markdown_parser import MarkdownElement


@dataclass
class WordStyle:
    """Word样式配置"""
    font_name: str = "微软雅黑"
    font_size: int = 12
    line_spacing: float = 1.15
    paragraph_spacing_before: int = 0
    paragraph_spacing_after: int = 6
    alignment: str = "left"
    bold: bool = False
    italic: bool = False
    color: Optional[str] = None


class WordGenerator:
    """Word文档生成器"""
    
    def __init__(self, template_path: Optional[str] = None, config: Optional[Dict] = None, style_handler=None):
        """初始化生成器
        
        Args:
            template_path: Word模板文件路径
            config: 生成器配置
            style_handler: 样式处理器（可选）
        """
        self.config = config or {}
        self.template_path = template_path
        self.style_handler = style_handler
        self.document = self._create_document()
        self.styles = self._setup_styles()
        
    def _create_document(self) -> Document:
        """创建Word文档"""
        if self.template_path and os.path.exists(self.template_path):
            return Document(self.template_path)
        else:
            return Document()
    
    def _setup_styles(self) -> Dict[str, WordStyle]:
        """设置文档样式"""
        styles = {
            'normal': WordStyle(
                font_name="微软雅黑",
                font_size=12,
                line_spacing=1.15
            ),
            'heading1': WordStyle(
                font_name="微软雅黑",
                font_size=18,
                bold=True,
                paragraph_spacing_before=12,
                paragraph_spacing_after=6
            ),
            'heading2': WordStyle(
                font_name="微软雅黑",
                font_size=16,
                bold=True,
                paragraph_spacing_before=10,
                paragraph_spacing_after=6
            ),
            'heading3': WordStyle(
                font_name="微软雅黑",
                font_size=14,
                bold=True,
                paragraph_spacing_before=8,
                paragraph_spacing_after=6
            ),
            'heading4': WordStyle(
                font_name="微软雅黑",
                font_size=13,
                bold=True,
                paragraph_spacing_before=6,
                paragraph_spacing_after=6
            ),
            'heading5': WordStyle(
                font_name="微软雅黑",
                font_size=12,
                bold=True,
                paragraph_spacing_before=6,
                paragraph_spacing_after=6
            ),
            'heading6': WordStyle(
                font_name="微软雅黑",
                font_size=12,
                bold=True,
                italic=True,
                paragraph_spacing_before=6,
                paragraph_spacing_after=6
            ),
            'code': WordStyle(
                font_name="Consolas",
                font_size=10,
                color="#333333"
            ),
            'quote': WordStyle(
                font_name="微软雅黑",
                font_size=12,
                italic=True,
                color="#666666"
            )
        }
        
        # 应用自定义样式配置
        if 'styles' in self.config:
            for style_name, style_config in self.config['styles'].items():
                if style_name in styles:
                    for attr, value in style_config.items():
                        setattr(styles[style_name], attr, value)
        
        return styles
    
    def generate(self, markdown_element: MarkdownElement, output_path: str) -> bool:
        """生成Word文档
        
        Args:
            markdown_element: 解析后的Markdown元素
            output_path: 输出文件路径
            
        Returns:
            是否生成成功
        """
        try:
            # 处理文档内容
            self._process_element(markdown_element)
            
            # 保存文档
            self.document.save(output_path)
            return True
            
        except Exception as e:
            print(f"生成Word文档失败: {e}")
            return False
    
    def generate_from_html(self, html_content: str, metadata: Dict[str, Any], output_path: str) -> bool:
        """从HTML内容生成Word文档（简化版本）
        
        Args:
            html_content: HTML内容
            metadata: 文档元数据
            output_path: 输出文件路径
            
        Returns:
            是否生成成功
        """
        try:
            from bs4 import BeautifulSoup
            
            # 解析HTML
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # 添加文档标题
            if metadata.get('title'):
                self._add_title(metadata['title'])
            
            # 处理HTML元素
            self._process_html_elements(soup)
            
            # 保存文档
            self.document.save(output_path)
            return True
            
        except Exception as e:
            print(f"生成Word文档失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _process_html_elements(self, soup):
        """处理HTML元素"""
        for element in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'pre', 'blockquote', 'ul', 'ol', 'table']):
            if element.name.startswith('h'):
                level = int(element.name[1])
                self.document.add_heading(element.get_text().strip(), level=level)
            elif element.name == 'p':
                paragraph = self.document.add_paragraph()
                self._process_paragraph_content(paragraph, element)
            elif element.name == 'pre':
                code_text = element.get_text().strip()
                code_paragraph = self.document.add_paragraph(code_text)
                self._apply_style(code_paragraph, self.styles['code'])
            elif element.name == 'blockquote':
                quote_text = element.get_text().strip()
                quote_paragraph = self.document.add_paragraph(quote_text)
                self._apply_style(quote_paragraph, self.styles['quote'])
                quote_paragraph.paragraph_format.left_indent = Inches(0.5)
            elif element.name in ['ul', 'ol']:
                self._process_html_list(element)
            elif element.name == 'table':
                # 在表格前添加一个空行
                self.document.add_paragraph()
                self._process_html_table(element)
                # 在表格后添加一个空行
                self.document.add_paragraph()
    
    def _process_paragraph_content(self, paragraph, element):
        """处理段落内容"""
        for content in element.contents:
            if hasattr(content, 'name'):
                if content.name == 'strong' or content.name == 'b':
                    run = paragraph.add_run(content.get_text())
                    run.bold = True
                elif content.name == 'em' or content.name == 'i':
                    run = paragraph.add_run(content.get_text())
                    run.italic = True
                elif content.name == 'code':
                    run = paragraph.add_run(content.get_text())
                    self._apply_code_style(run)
                elif content.name == 'a':
                    text = content.get_text()
                    url = content.get('href', '#')
                    self._add_hyperlink(paragraph, url, text)
                else:
                    paragraph.add_run(content.get_text())
            else:
                paragraph.add_run(str(content))
    
    def _process_html_list(self, element, indent_level: int = 0):
        """处理HTML列表
        
        Args:
            element: HTML列表元素
            indent_level: 缩进级别（用于嵌套列表）
        """
        is_ordered = element.name == 'ol'
        
        # 计算左缩进：基础缩进 + 嵌套缩进
        # 基础缩进：0.5 英寸（约 1.27 厘米），使列表比段落有更多缩进
        # 嵌套缩进：每级嵌套增加 0.5 英寸
        base_indent = Inches(0.5)
        nested_indent = Inches(0.5 * indent_level)
        total_indent = base_indent + nested_indent
        
        for li in element.find_all('li', recursive=False):
            # 根据列表类型选择样式
            if is_ordered:
                # 有序列表使用 List Number 样式
                paragraph = self.document.add_paragraph(style='List Number')
            else:
                # 无序列表使用 List Bullet 样式
                paragraph = self.document.add_paragraph(style='List Bullet')
            
            # 设置列表项的左缩进，使其比普通段落有更多缩进
            paragraph.paragraph_format.left_indent = total_indent
            
            # 处理列表项内容，包括格式化文本（粗体、斜体、代码、链接等）
            self._process_paragraph_content(paragraph, li)
            
            # 处理嵌套列表
            nested_lists = li.find_all(['ul', 'ol'], recursive=False)
            for nested_list in nested_lists:
                self._process_html_list(nested_list, indent_level + 1)
    
    def _process_html_table(self, element):
        """处理HTML表格"""
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        
        rows = element.find_all('tr')
        if not rows:
            return
        
        # 计算列数
        max_cols = max(len(row.find_all(['td', 'th'])) for row in rows)
        
        # 创建表格
        table = self.document.add_table(rows=len(rows), cols=max_cols)
        table.style = 'Table Grid'
        
        # 禁用自动调整列宽，以便手动控制列宽
        table.autofit = False
        
        # 设置表格属性：对齐方式、文字环绕、单元格边距
        self._setup_table_properties(table)
        
        # 计算每列的最大内容长度（用于智能列宽分配）
        column_max_lengths = [0] * max_cols
        for row in rows:
            cells = row.find_all(['td', 'th'])
            for j, cell in enumerate(cells):
                if j < max_cols:
                    cell_text = cell.get_text().strip()
                    # 计算文本长度（中文字符按2个字符计算）
                    text_length = sum(2 if ord(c) > 127 else 1 for c in cell_text)
                    column_max_lengths[j] = max(column_max_lengths[j], text_length)
        
        # 填充数据并设置样式
        for i, row in enumerate(rows):
            cells = row.find_all(['td', 'th'])
            for j, cell in enumerate(cells):
                if j < max_cols:
                    word_cell = table.cell(i, j)
                    word_cell.text = cell.get_text().strip()
                    
                    # 设置单元格对齐方式为居中
                    for paragraph in word_cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        # 设置字体为微软雅黑，小五（9号）
                        for run in paragraph.runs:
                            run.font.name = '微软雅黑'
                            run.font.size = Pt(9)
        
        # 智能调整列宽（设置表格整体宽度最小值和最大值）
        self._adjust_table_column_widths(table, column_max_lengths, min_table_width_cm=8.0, max_table_width_cm=16.0)
    
    def _process_element(self, element: MarkdownElement):
        """处理Markdown元素"""
        if element.element_type == 'document':
            self._process_document(element)
        elif element.element_type.startswith('heading'):
            self._process_heading(element)
        elif element.element_type == 'paragraph':
            self._process_paragraph(element)
        elif element.element_type == 'code_block':
            self._process_code_block(element)
        elif element.element_type == 'table':
            self._process_table(element)
        elif element.element_type == 'image':
            self._process_image(element)
        elif element.element_type == 'list':
            self._process_list(element)
        elif element.element_type == 'quote':
            self._process_quote(element)
        
        # 处理子元素
        for child in element.children:
            self._process_element(child)
    
    def _process_document(self, element: MarkdownElement):
        """处理文档根元素"""
        # 不再自动添加标题，标题应该由 Markdown 中的 H1 标题处理
        # if 'title' in element.attributes:
        #     self._add_title(element.attributes['title'])
        pass
    
    def _process_heading(self, element: MarkdownElement):
        """处理标题"""
        level = int(element.element_type.replace('heading', ''))
        style_name = f'heading{level}'
        
        paragraph = self.document.add_heading(element.content, level=level)
        
        # 优先使用 StyleHandler 的样式
        if self.style_handler:
            style = self.style_handler.get_style(style_name)
            if style:
                self._apply_style_handler_style(paragraph, style, is_heading=True)
                return
        
        # 否则使用默认样式
        self._apply_style(paragraph, self.styles.get(style_name, self.styles['normal']))
    
    def _process_paragraph(self, element: MarkdownElement):
        """处理段落"""
        paragraph = self.document.add_paragraph()
        
        # 处理段落内容，包括格式化文本
        self._process_formatted_text(paragraph, element.content)
        
        # 优先使用 StyleHandler 的样式
        if self.style_handler:
            style = self.style_handler.get_style('normal')
            if style:
                self._apply_style_handler_style(paragraph, style, is_heading=False)
                return
        
        # 否则使用默认样式
        self._apply_style(paragraph, self.styles['normal'])
    
    def _process_formatted_text(self, paragraph, text: str):
        """处理格式化文本"""
        # 处理粗体、斜体、代码等格式
        parts = self._split_formatted_text(text)
        
        for part in parts:
            if part['type'] == 'bold':
                run = paragraph.add_run(part['text'])
                run.bold = True
            elif part['type'] == 'italic':
                run = paragraph.add_run(part['text'])
                run.italic = True
            elif part['type'] == 'code':
                run = paragraph.add_run(part['text'])
                self._apply_code_style(run)
            elif part['type'] == 'link':
                self._add_hyperlink(paragraph, part['url'], part['text'])
            else:
                paragraph.add_run(part['text'])
    
    def _split_formatted_text(self, text: str) -> List[Dict[str, str]]:
        """分割格式化文本"""
        if not text:
            # 如果文本为空，返回空列表（不添加任何内容）
            return []
        
        parts = []
        current_pos = 0
        
        # 匹配各种格式的正则表达式
        patterns = [
            (r'\*\*([^*]+)\*\*', 'bold'),
            (r'\*([^*]+)\*', 'italic'),
            (r'`([^`]+)`', 'code'),
            (r'\[([^\]]+)\]\(([^)]+)\)', 'link')
        ]
        
        for pattern, format_type in patterns:
            for match in re.finditer(pattern, text):
                # 添加匹配前的普通文本
                if match.start() > current_pos:
                    parts.append({
                        'type': 'normal',
                        'text': text[current_pos:match.start()]
                    })
                
                # 添加格式化文本
                if format_type == 'link':
                    parts.append({
                        'type': 'link',
                        'text': match.group(1),
                        'url': match.group(2)
                    })
                else:
                    parts.append({
                        'type': format_type,
                        'text': match.group(1)
                    })
                
                current_pos = match.end()
        
        # 添加剩余的普通文本
        if current_pos < len(text):
            parts.append({
                'type': 'normal',
                'text': text[current_pos:]
            })
        
        # 如果没有找到任何部分，返回整个文本作为普通文本
        if not parts:
            parts.append({
                'type': 'normal',
                'text': text
            })
        
        return parts
    
    def _process_code_block(self, element: MarkdownElement):
        """处理代码块"""
        # 添加代码块标题（如果有语言标识）
        if 'language' in element.attributes:
            lang_paragraph = self.document.add_paragraph(f"代码 ({element.attributes['language']})")
            lang_paragraph.style = 'Caption'
        
        # 添加代码内容
        code_paragraph = self.document.add_paragraph(element.content)
        self._apply_style(code_paragraph, self.styles['code'])
        
        # 设置代码块背景色
        self._set_paragraph_background(code_paragraph, '#f8f8f8')
    
    def _process_table(self, element: MarkdownElement):
        """处理表格"""
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        
        if 'rows' not in element.attributes or 'cols' not in element.attributes:
            return
        
        # 在表格前添加一个空行
        self.document.add_paragraph()
        
        rows = element.attributes['rows']
        cols = element.attributes['cols']
        
        table = self.document.add_table(rows=rows, cols=cols)
        table.style = 'Table Grid'
        
        # 禁用自动调整列宽，以便手动控制列宽
        table.autofit = False
        
        # 设置表格属性：对齐方式、文字环绕、单元格边距
        self._setup_table_properties(table)
        
        # 填充表格数据并设置样式
        if 'data' in element.attributes:
            data = element.attributes['data']
            
            # 计算每列的最大内容长度（用于智能列宽分配）
            column_max_lengths = [0] * cols
            for i, row_data in enumerate(data):
                if i < rows:
                    for j, cell_data in enumerate(row_data):
                        if j < cols:
                            cell_text = str(cell_data)
                            # 计算文本长度（中文字符按2个字符计算）
                            text_length = sum(2 if ord(c) > 127 else 1 for c in cell_text)
                            column_max_lengths[j] = max(column_max_lengths[j], text_length)
            
            # 填充数据并设置样式
            for i, row_data in enumerate(data):
                if i < rows:
                    for j, cell_data in enumerate(row_data):
                        if j < cols:
                            cell = table.cell(i, j)
                            cell.text = str(cell_data)
                            
                            # 设置单元格对齐方式为居中
                            for paragraph in cell.paragraphs:
                                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                
                                # 设置字体为微软雅黑，小五（9号）
                                for run in paragraph.runs:
                                    run.font.name = '微软雅黑'
                                    run.font.size = Pt(9)
            
            # 智能调整列宽（设置表格整体宽度最小值和最大值）
            self._adjust_table_column_widths(table, column_max_lengths, min_table_width_cm=8.0, max_table_width_cm=16.0)
        
        # 在表格后添加一个空行
        self.document.add_paragraph()
    
    def _setup_table_properties(self, table):
        """设置表格属性：对齐方式、文字环绕、单元格边距
        
        Args:
            table: Word表格对象
        """
        # 1. 设置表格对齐方式为居中
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # 2. 设置文字环绕方式为环绕
        tbl = table._tbl  # 获取底层的XML元素
        
        # 查找或创建tblPr元素
        tblPr_elements = tbl.findall(qn('w:tblPr'))
        if not tblPr_elements:
            tblPr = OxmlElement('w:tblPr')
            # 将tblPr插入到tbl的第一个位置
            tbl.insert(0, tblPr)
        else:
            tblPr = tblPr_elements[0]
        
        # 查找或创建tblpPr元素
        tblpPr_elements = tblPr.findall(qn('w:tblpPr'))
        if not tblpPr_elements:
            tblpPr = OxmlElement('w:tblpPr')
            tblPr.append(tblpPr)
        else:
            tblpPr = tblpPr_elements[0]
        
        # 设置文字环绕为环绕
        tblpPr.set(qn('w:wrap'), 'around')
        
        # 3. 设置单元格边距：上下均为0.15cm
        # 查找或创建tblCellMar元素
        tblCellMar_elements = tblPr.findall(qn('w:tblCellMar'))
        if not tblCellMar_elements:
            tblCellMar = OxmlElement('w:tblCellMar')
            tblPr.append(tblCellMar)
        else:
            tblCellMar = tblCellMar_elements[0]
        
        # 清除现有的边距设置（如果存在）
        for margin_name in ['w:top', 'w:bottom', 'w:left', 'w:right']:
            existing = tblCellMar.findall(qn(margin_name))
            for elem in existing:
                tblCellMar.remove(elem)
        
        # 设置上下边距为0.15cm
        top_margin = OxmlElement('w:top')
        top_margin.set(qn('w:w'), str(int(0.15 * 567)))  # 0.15cm转换为twips (1cm = 567 twips)
        top_margin.set(qn('w:type'), 'dxa')
        tblCellMar.append(top_margin)
        
        bottom_margin = OxmlElement('w:bottom')
        bottom_margin.set(qn('w:w'), str(int(0.15 * 567)))
        bottom_margin.set(qn('w:type'), 'dxa')
        tblCellMar.append(bottom_margin)
        
        # 设置左右边距为默认值（如果需要）
        left_margin = OxmlElement('w:left')
        left_margin.set(qn('w:w'), str(int(0.19 * 567)))  # 默认0.19cm
        left_margin.set(qn('w:type'), 'dxa')
        tblCellMar.append(left_margin)
        
        right_margin = OxmlElement('w:right')
        right_margin.set(qn('w:w'), str(int(0.19 * 567)))  # 默认0.19cm
        right_margin.set(qn('w:type'), 'dxa')
        tblCellMar.append(right_margin)
    
    def _adjust_table_column_widths(self, table, column_max_lengths, min_table_width_cm=8.0, max_table_width_cm=16.0):
        """根据内容长度智能调整表格列宽，并设置表格整体宽度的最小值和最大值
        
        Args:
            table: Word表格对象
            column_max_lengths: 每列的最大内容长度列表
            min_table_width_cm: 表格整体最小宽度（厘米），默认8.0cm
            max_table_width_cm: 表格整体最大宽度（厘米），默认16.0cm
        """
        if not column_max_lengths or sum(column_max_lengths) == 0:
            # 如果没有内容，设置默认宽度
            default_width_cm = (min_table_width_cm + max_table_width_cm) / 2
            if len(table.columns) > 0:
                col_width_cm = default_width_cm / len(table.columns)
                col_width = Cm(col_width_cm)
            else:
                col_width = Cm(default_width_cm)
            for col in table.columns:
                col.width = col_width
            return
        
        # 设置最小和最大列宽（厘米）
        min_col_width_cm = 1.2  # 最小列宽1.2cm（用于序号等短内容）
        max_col_width_cm = 6.0  # 最大列宽6.0cm（用于长描述）
        
        # 计算所有列的最大内容长度的中位数作为基准
        valid_lengths = [length for length in column_max_lengths if length > 0]
        if not valid_lengths:
            # 如果没有有效长度，使用默认宽度
            default_width_cm = (min_col_width_cm + max_col_width_cm) / 2
            column_widths_cm = [default_width_cm] * len(column_max_lengths)
        else:
            # 计算中位数
            sorted_lengths = sorted(valid_lengths)
            median_length = sorted_lengths[len(sorted_lengths) // 2]
            
            # 如果中位数为0或太小，使用平均值作为基准
            if median_length < 3:
                median_length = sum(valid_lengths) / len(valid_lengths) if valid_lengths else 10
            
            # 根据中位数计算基准宽度（中位数对应的列宽设为中等宽度）
            # 假设中位数对应的列宽为3.0cm（表格总宽度在8-16cm之间，5列的话平均约3cm）
            base_width_cm = 3.0
            char_width_cm = base_width_cm / median_length  # 每个字符对应的宽度
            
            # 根据内容计算每列需要的宽度
            column_widths_cm = []
            for max_length in column_max_lengths:
                if max_length > 0:
                    # 对于非常短的内容（<=3个字符），使用最小宽度
                    if max_length <= 3:
                        col_width_cm = min_col_width_cm
                    else:
                        # 根据字符数与中位数的比例计算列宽
                        # 如果内容长度等于中位数，则宽度为base_width_cm
                        # 如果内容长度是中位数的2倍，则宽度约为base_width_cm的2倍（但不超过最大值）
                        col_width_cm = max_length * char_width_cm
                        # 确保在最小和最大宽度之间
                        col_width_cm = max(min_col_width_cm, min(col_width_cm, max_col_width_cm))
                else:
                    # 如果列没有内容，使用最小宽度
                    col_width_cm = min_col_width_cm
                column_widths_cm.append(col_width_cm)
        
        # 计算表格总宽度
        total_width_cm = sum(column_widths_cm)
        
        # 如果总宽度超出范围，按比例缩放所有列
        if total_width_cm < min_table_width_cm:
            # 如果总宽度小于最小值，按比例放大
            scale_factor = min_table_width_cm / total_width_cm
            column_widths_cm = [w * scale_factor for w in column_widths_cm]
            total_width_cm = min_table_width_cm
        elif total_width_cm > max_table_width_cm:
            # 如果总宽度大于最大值，按比例缩小
            scale_factor = max_table_width_cm / total_width_cm
            column_widths_cm = [w * scale_factor for w in column_widths_cm]
            total_width_cm = max_table_width_cm
        
        # 确保每列宽度在合理范围内（缩放后可能超出范围）
        for j, width_cm in enumerate(column_widths_cm):
            width_cm = max(min_col_width_cm, min(width_cm, max_col_width_cm))
            column_widths_cm[j] = width_cm
        
        # 应用列宽 - 同时设置列和单元格的宽度以确保生效
        for j, width_cm in enumerate(column_widths_cm):
            col_width = Cm(width_cm)
            # 设置列的宽度
            table.columns[j].width = col_width
            # 同时设置该列所有单元格的宽度，确保生效
            for cell in table.columns[j].cells:
                cell.width = col_width
    
    def _process_image(self, element: MarkdownElement):
        """处理图片"""
        if 'src' not in element.attributes:
            return
        
        src = element.attributes['src']
        alt = element.attributes.get('alt', '')
        
        try:
            # 如果是本地文件路径
            if os.path.exists(src):
                paragraph = self.document.add_paragraph()
                run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
                run.add_picture(src, width=Inches(6))
                
                # 添加图片说明
                if alt:
                    caption = self.document.add_paragraph(f"图: {alt}")
                    caption.style = 'Caption'
                    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                # 网络图片或无效路径，添加占位符
                placeholder = self.document.add_paragraph(f"[图片: {alt or src}]")
                placeholder.style = 'Caption'
                
        except Exception as e:
            print(f"处理图片失败 {src}: {e}")
            # 添加错误占位符
            error_placeholder = self.document.add_paragraph(f"[图片加载失败: {alt or src}]")
            error_placeholder.style = 'Caption'
    
    def _process_list(self, element: MarkdownElement, indent_level: int = 0):
        """处理列表
        
        Args:
            element: 列表元素
            indent_level: 缩进级别（用于嵌套列表）
        """
        list_type = element.attributes.get('type', 'unordered')
        
        # 计算左缩进：基础缩进 + 嵌套缩进
        # 基础缩进：0.5 英寸（约 1.27 厘米），使列表比段落有更多缩进
        # 嵌套缩进：每级嵌套增加 0.5 英寸
        base_indent = Inches(0.5)
        nested_indent = Inches(0.5 * indent_level)
        total_indent = base_indent + nested_indent
        
        for item in element.children:
            # 根据列表类型选择样式
            if list_type == 'ordered':
                # 有序列表使用 List Number 样式
                paragraph = self.document.add_paragraph(style='List Number')
            else:
                # 无序列表使用 List Bullet 样式
                paragraph = self.document.add_paragraph(style='List Bullet')
            
            # 设置列表项的左缩进，使其比普通段落有更多缩进
            paragraph.paragraph_format.left_indent = total_indent
            
            # 处理列表项内容，包括格式化文本（粗体、斜体、代码、链接等）
            if item.content:
                self._process_formatted_text(paragraph, item.content)
            
            # 处理列表项的子元素（如嵌套列表、段落等）
            for child in item.children:
                if child.element_type == 'list':
                    # 嵌套列表：递归处理，缩进级别+1
                    self._process_list(child, indent_level + 1)
                elif child.element_type == 'paragraph':
                    # 列表项内的段落：作为列表项的一部分处理
                    self._process_formatted_text(paragraph, child.content)
    
    def _process_quote(self, element: MarkdownElement):
        """处理引用"""
        paragraph = self.document.add_paragraph(element.content)
        self._apply_style(paragraph, self.styles['quote'])
        
        # 设置引用样式
        paragraph.paragraph_format.left_indent = Inches(0.5)
        self._set_paragraph_border(paragraph, 'left')
    
    def _apply_style_handler_style(self, paragraph, element_style, is_heading=False):
        """应用 StyleHandler 的样式到段落
        
        Args:
            paragraph: Word段落对象
            element_style: 元素样式
            is_heading: 是否为标题，标题不应用首行缩进
        """
        from docx.enum.text import WD_LINE_SPACING
        
        # 应用字体样式
        for run in paragraph.runs:
            font = run.font
            font.name = element_style.font.name
            font.size = Pt(element_style.font.size)
            font.bold = element_style.font.bold
            font.italic = element_style.font.italic
            
            if element_style.font.color:
                color_hex = element_style.font.color.replace('#', '')
                if len(color_hex) == 6:
                    r = int(color_hex[0:2], 16)
                    g = int(color_hex[2:4], 16)
                    b = int(color_hex[4:6], 16)
                    font.color.rgb = RGBColor(r, g, b)
        
        # 应用段落样式
        paragraph_format = paragraph.paragraph_format
        
        # 对齐方式
        alignment_map = {
            'left': WD_ALIGN_PARAGRAPH.LEFT,
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT,
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        paragraph.alignment = alignment_map.get(element_style.paragraph.alignment, WD_ALIGN_PARAGRAPH.LEFT)
        
        # 行距处理：如果 line_spacing >= 20，认为是固定值（磅），否则认为是倍数
        if element_style.paragraph.line_spacing >= 20:
            # 固定值行距（磅）
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            paragraph_format.line_spacing = Pt(element_style.paragraph.line_spacing)
        else:
            # 倍数行距
            paragraph_format.line_spacing = element_style.paragraph.line_spacing
        
        # 段前段后间距
        paragraph_format.space_before = Pt(element_style.paragraph.space_before)
        paragraph_format.space_after = Pt(element_style.paragraph.space_after)
        
        # 缩进
        # 注意：配置中的值单位是厘米，需要转换为英寸
        # 1厘米 = 1/2.54 英寸
        paragraph_format.left_indent = Inches(element_style.paragraph.left_indent / 2.54)
        paragraph_format.right_indent = Inches(element_style.paragraph.right_indent / 2.54)
        
        # 首行缩进：标题不应用首行缩进，只有正文段落才应用
        if is_heading:
            # 标题不应用首行缩进
            paragraph_format.first_line_indent = Inches(0)
        elif element_style.paragraph.first_line_indent == 0 and element_style.font.size:
            # 正文段落：如果配置值为0，则根据字体大小动态计算两个字符的宽度
            # 两个字符宽度 = 字体大小 × 2（磅），转换为英寸
            # 1磅 = 1/72 英寸
            first_line_indent_inches = (element_style.font.size * 2) / 72.0
            paragraph_format.first_line_indent = Inches(first_line_indent_inches)
        else:
            # 使用配置值（厘米转英寸）
            paragraph_format.first_line_indent = Inches(element_style.paragraph.first_line_indent / 2.54)
        
        # 分页控制
        paragraph_format.keep_together = element_style.paragraph.keep_together
        paragraph_format.keep_with_next = element_style.paragraph.keep_with_next
        paragraph_format.page_break_before = element_style.paragraph.page_break_before
    
    def _apply_style(self, paragraph, style: WordStyle):
        """应用样式到段落"""
        from docx.enum.text import WD_LINE_SPACING
        
        # 设置字体
        for run in paragraph.runs:
            font = run.font
            font.name = style.font_name
            font.size = Pt(style.font_size)
            font.bold = style.bold
            font.italic = style.italic
            
            if style.color:
                font.color.rgb = RGBColor.from_string(style.color.replace('#', ''))
        
        # 设置段落格式
        paragraph_format = paragraph.paragraph_format
        
        # 行距处理：如果 line_spacing >= 20，认为是固定值（磅），否则认为是倍数
        if style.line_spacing >= 20:
            # 固定值行距（磅）
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            paragraph_format.line_spacing = Pt(style.line_spacing)
        else:
            # 倍数行距
            paragraph_format.line_spacing = style.line_spacing
        
        paragraph_format.space_before = Pt(style.paragraph_spacing_before)
        paragraph_format.space_after = Pt(style.paragraph_spacing_after)
        
        # 设置对齐方式
        if style.alignment == 'center':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif style.alignment == 'right':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif style.alignment == 'justify':
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        else:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    def _apply_code_style(self, run):
        """应用代码样式"""
        font = run.font
        font.name = 'Consolas'
        font.size = Pt(10)
        font.color.rgb = RGBColor(51, 51, 51)
    
    def _add_hyperlink(self, paragraph, url: str, text: str):
        """添加超链接"""
        # 创建超链接
        part = paragraph.part
        r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
        
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('r:id'), r_id)
        
        new_run = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        
        # 设置超链接样式
        color = OxmlElement('w:color')
        color.set(qn('w:val'), '0000FF')
        rPr.append(color)
        
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)
        
        new_run.append(rPr)
        new_run.text = text
        hyperlink.append(new_run)
        
        paragraph._p.append(hyperlink)
    
    def _set_paragraph_background(self, paragraph, color: str):
        """设置段落背景色"""
        shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color.replace("#", "")}" />')
        paragraph._p.get_or_add_pPr().append(shading_elm)
    
    def _set_paragraph_border(self, paragraph, side: str = 'left'):
        """设置段落边框"""
        pPr = paragraph._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        
        border = OxmlElement(f'w:{side}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '1')
        border.set(qn('w:color'), 'CCCCCC')
        
        pBdr.append(border)
        pPr.append(pBdr)
    
    def _add_title(self, title: str):
        """添加文档标题"""
        title_paragraph = self.document.add_heading(title, level=0)
        title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    def add_page_break(self):
        """添加分页符"""
        paragraph = self.document.add_paragraph()
        run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
        run.add_break(WD_BREAK.PAGE)
    
    def set_page_margins(self, top: float = 1.0, bottom: float = 1.0, 
                        left: float = 1.0, right: float = 1.0):
        """设置页面边距（英寸）"""
        sections = self.document.sections
        for section in sections:
            section.top_margin = Inches(top)
            section.bottom_margin = Inches(bottom)
            section.left_margin = Inches(left)
            section.right_margin = Inches(right)