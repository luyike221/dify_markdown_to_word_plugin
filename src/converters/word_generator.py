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
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.enum.style import WD_STYLE_TYPE
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
    
    def __init__(self, template_path: Optional[str] = None, config: Optional[Dict] = None):
        """初始化生成器
        
        Args:
            template_path: Word模板文件路径
            config: 生成器配置
        """
        self.config = config or {}
        self.template_path = template_path
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
                self._process_html_table(element)
    
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
    
    def _process_html_list(self, element):
        """处理HTML列表"""
        is_ordered = element.name == 'ol'
        for li in element.find_all('li', recursive=False):
            text = li.get_text().strip()
            if is_ordered:
                self.document.add_paragraph(text, style='List Number')
            else:
                self.document.add_paragraph(text, style='List Bullet')
    
    def _process_html_table(self, element):
        """处理HTML表格"""
        rows = element.find_all('tr')
        if not rows:
            return
        
        # 计算列数
        max_cols = max(len(row.find_all(['td', 'th'])) for row in rows)
        
        # 创建表格
        table = self.document.add_table(rows=len(rows), cols=max_cols)
        table.style = 'Table Grid'
        
        # 填充数据
        for i, row in enumerate(rows):
            cells = row.find_all(['td', 'th'])
            for j, cell in enumerate(cells):
                if j < max_cols:
                    table.cell(i, j).text = cell.get_text().strip()
    
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
        # 设置文档属性
        if 'title' in element.attributes:
            self._add_title(element.attributes['title'])
    
    def _process_heading(self, element: MarkdownElement):
        """处理标题"""
        level = int(element.element_type.replace('heading', ''))
        style_name = f'heading{level}'
        
        paragraph = self.document.add_heading(element.content, level=level)
        self._apply_style(paragraph, self.styles.get(style_name, self.styles['normal']))
    
    def _process_paragraph(self, element: MarkdownElement):
        """处理段落"""
        paragraph = self.document.add_paragraph()
        
        # 处理段落内容，包括格式化文本
        self._process_formatted_text(paragraph, element.content)
        
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
        if 'rows' not in element.attributes or 'cols' not in element.attributes:
            return
        
        rows = element.attributes['rows']
        cols = element.attributes['cols']
        
        table = self.document.add_table(rows=rows, cols=cols)
        table.style = 'Table Grid'
        
        # 填充表格数据
        if 'data' in element.attributes:
            data = element.attributes['data']
            for i, row_data in enumerate(data):
                if i < rows:
                    for j, cell_data in enumerate(row_data):
                        if j < cols:
                            table.cell(i, j).text = str(cell_data)
    
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
    
    def _process_list(self, element: MarkdownElement):
        """处理列表"""
        list_type = element.attributes.get('type', 'unordered')
        
        for item in element.children:
            if list_type == 'ordered':
                paragraph = self.document.add_paragraph(item.content, style='List Number')
            else:
                paragraph = self.document.add_paragraph(item.content, style='List Bullet')
    
    def _process_quote(self, element: MarkdownElement):
        """处理引用"""
        paragraph = self.document.add_paragraph(element.content)
        self._apply_style(paragraph, self.styles['quote'])
        
        # 设置引用样式
        paragraph.paragraph_format.left_indent = Inches(0.5)
        self._set_paragraph_border(paragraph, 'left')
    
    def _apply_style(self, paragraph, style: WordStyle):
        """应用样式到段落"""
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