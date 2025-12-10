#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown解析器模块
负责解析Markdown文档，构建语法树，处理各种Markdown元素
"""

import re
import markdown
from markdown.extensions import codehilite, tables, toc
from typing import Dict, List, Any, Optional
from dataclasses import dataclass


@dataclass
class MarkdownElement:
    """Markdown元素基类"""
    element_type: str
    content: str
    attributes: Dict[str, Any] = None
    children: List['MarkdownElement'] = None

    def __post_init__(self):
        if self.attributes is None:
            self.attributes = {}
        if self.children is None:
            self.children = []


class MarkdownParser:
    """Markdown解析器"""
    
    def __init__(self, config: Optional[Dict] = None):
        """初始化解析器
        
        Args:
            config: 解析器配置
        """
        self.config = config or {}
        self.markdown_instance = self._create_markdown_instance()
        
    def _create_markdown_instance(self) -> markdown.Markdown:
        """创建Markdown实例"""
        extensions = [
            'markdown.extensions.tables',
            'markdown.extensions.codehilite',
            'markdown.extensions.toc',
            'markdown.extensions.fenced_code',
            'markdown.extensions.attr_list',
            'markdown.extensions.def_list',
            'markdown.extensions.footnotes',
            'markdown.extensions.md_in_html'
        ]
        
        extension_configs = {
            'codehilite': {
                'css_class': 'highlight',
                'use_pygments': True
            },
            'toc': {
                'permalink': True
            }
        }
        
        return markdown.Markdown(
            extensions=extensions,
            extension_configs=extension_configs
        )
    
    def parse(self, markdown_text: str) -> MarkdownElement:
        """解析Markdown文本
        
        Args:
            markdown_text: Markdown文本内容
            
        Returns:
            解析后的文档树
        """
        # 预处理
        processed_text = self._preprocess(markdown_text)
        
        # 解析为HTML
        html_content = self.markdown_instance.convert(processed_text)
        
        # 构建文档树
        document = self._build_document_tree(html_content, processed_text)
        
        return document
    
    def _preprocess(self, text: str) -> str:
        """预处理Markdown文本"""
        # 处理转义的换行符（兼容 \\n 格式）
        # 将字面上的 \n 转换为真正的换行符
        text = text.replace('\\n', '\n')
        
        # 标准化换行符
        text = text.replace('\r\n', '\n').replace('\r', '\n')
        
        # 处理数学公式
        text = self._process_math_formulas(text)
        
        # 处理图片链接
        text = self._process_images(text)
        
        # 确保列表前有空行，以便正确解析
        text = self._ensure_blank_line_before_lists(text)
        
        return text
    
    def _process_math_formulas(self, text: str) -> str:
        """处理数学公式"""
        # 行内公式 $...$
        inline_pattern = r'\$([^$]+)\$'
        text = re.sub(inline_pattern, r'<math-inline>\1</math-inline>', text)
        
        # 块级公式 $$...$$
        block_pattern = r'\$\$([^$]+)\$\$'
        text = re.sub(block_pattern, r'<math-block>\1</math-block>', text, flags=re.MULTILINE | re.DOTALL)
        
        return text
    
    def _process_images(self, text: str) -> str:
        """处理图片链接"""
        # 提取图片信息并标记
        image_pattern = r'!\[([^\]]*)\]\(([^)]+)\)'
        
        def replace_image(match):
            alt_text = match.group(1)
            src = match.group(2)
            return f'<img-placeholder alt="{alt_text}" src="{src}"></img-placeholder>'
        
        return re.sub(image_pattern, replace_image, text)
    
    def _ensure_blank_line_before_lists(self, text: str) -> str:
        """确保列表前有空行，以便正确解析
        
        Args:
            text: Markdown文本
            
        Returns:
            处理后的文本
        """
        lines = text.split('\n')
        result_lines = []
        
        for i, line in enumerate(lines):
            # 检查是否是列表项（有序或无序）
            is_list_item = (
                re.match(r'^\s*(\d+\.|\*|\-|\+)\s+', line) is not None
            )
            
            if is_list_item and i > 0:
                # 检查前一行是否为空行
                prev_line = lines[i - 1].strip() if i > 0 else ''
                # 如果前一行不为空且不是列表项，则添加空行
                if prev_line and not re.match(r'^\s*(\d+\.|\*|\-|\+)\s+', lines[i - 1]):
                    result_lines.append('')
            
            result_lines.append(line)
        
        return '\n'.join(result_lines)
    
    def _build_document_tree(self, html_content: str, original_text: str) -> MarkdownElement:
        """构建文档树"""
        try:
            from bs4 import BeautifulSoup
            
            # 解析HTML内容
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # 创建文档根元素
            document = MarkdownElement(
                element_type='document',
                content='',
                attributes={'original_text': original_text}
            )
            
            # 不再自动提取和添加标题，标题应该由 Markdown 中的 H1 标题处理
            # title = self._extract_title(original_text)
            # if title:
            #     document.attributes['title'] = title
            
            # 解析HTML元素并构建文档树
            self._parse_html_elements(soup, document)
            
            return document
        except ImportError:
            # 如果没有BeautifulSoup，使用简单方法
            # 直接从原始Markdown文本解析
            return self._build_document_tree_from_markdown(original_text)
        except Exception as e:
            print(f"解析HTML失败，使用备用方法: {e}")
            return self._build_document_tree_from_markdown(original_text)
    
    def _parse_html_elements(self, soup, parent: MarkdownElement):
        """解析HTML元素并添加到文档树"""
        for element in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'pre', 'blockquote', 'ul', 'ol', 'table', 'img']):
            if element.name.startswith('h'):
                # 处理标题
                level = int(element.name[1])
                heading = MarkdownElement(
                    element_type=f'heading{level}',
                    content=element.get_text().strip(),
                    attributes={}
                )
                parent.children.append(heading)
            elif element.name == 'p':
                # 处理段落
                text = element.get_text().strip()
                if text:  # 只添加非空段落
                    paragraph = MarkdownElement(
                        element_type='paragraph',
                        content=text,
                        attributes={}
                    )
                    parent.children.append(paragraph)
            elif element.name == 'pre':
                # 处理代码块
                code_text = element.get_text().strip()
                code_element = element.find('code')
                language = code_element.get('class', [''])[0].replace('language-', '') if code_element else ''
                code_block = MarkdownElement(
                    element_type='code_block',
                    content=code_text,
                    attributes={'language': language} if language else {}
                )
                parent.children.append(code_block)
            elif element.name == 'blockquote':
                # 处理引用
                quote_text = element.get_text().strip()
                quote = MarkdownElement(
                    element_type='quote',
                    content=quote_text,
                    attributes={}
                )
                parent.children.append(quote)
            elif element.name in ['ul', 'ol']:
                # 处理列表
                list_type = 'ordered' if element.name == 'ol' else 'unordered'
                list_element = MarkdownElement(
                    element_type='list',
                    content='',
                    attributes={'type': list_type}
                )
                for li in element.find_all('li', recursive=False):
                    item_text = li.get_text().strip()
                    if item_text:
                        list_item = MarkdownElement(
                            element_type='list_item',
                            content=item_text,
                            attributes={}
                        )
                        list_element.children.append(list_item)
                if list_element.children:
                    parent.children.append(list_element)
            elif element.name == 'table':
                # 处理表格
                rows = element.find_all('tr')
                if rows:
                    table_data = []
                    for row in rows:
                        cells = row.find_all(['td', 'th'])
                        row_data = [cell.get_text().strip() for cell in cells]
                        if row_data:
                            table_data.append(row_data)
                    
                    if table_data:
                        table = MarkdownElement(
                            element_type='table',
                            content='',
                            attributes={
                                'rows': len(table_data),
                                'cols': max(len(row) for row in table_data) if table_data else 0,
                                'data': table_data
                            }
                        )
                        parent.children.append(table)
            elif element.name == 'img':
                # 处理图片
                src = element.get('src', '')
                alt = element.get('alt', '')
                if src:
                    image = MarkdownElement(
                        element_type='image',
                        content=alt,
                        attributes={'src': src, 'alt': alt}
                    )
                    parent.children.append(image)
    
    def _build_document_tree_from_markdown(self, markdown_text: str) -> MarkdownElement:
        """从Markdown文本直接构建文档树（备用方法）"""
        document = MarkdownElement(
            element_type='document',
            content='',
            attributes={'original_text': markdown_text}
        )
        
        # 不再自动提取和添加标题，标题应该由 Markdown 中的 H1 标题处理
        # title = self._extract_title(markdown_text)
        # if title:
        #     document.attributes['title'] = title
        
        lines = markdown_text.split('\n')
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            
            # 处理标题
            if line.startswith('#'):
                level = len(line) - len(line.lstrip('#'))
                heading_text = line.lstrip('#').strip()
                if heading_text and level <= 6:
                    heading = MarkdownElement(
                        element_type=f'heading{level}',
                        content=heading_text,
                        attributes={}
                    )
                    document.children.append(heading)
            
            # 处理段落
            elif line and not line.startswith('|') and not line.startswith('```') and not line.startswith('>'):
                # 收集连续的非空行作为段落
                paragraph_lines = [line]
                i += 1
                while i < len(lines) and lines[i].strip() and not lines[i].strip().startswith(('#', '|', '```', '>', '-', '*')):
                    paragraph_lines.append(lines[i].strip())
                    i += 1
                i -= 1  # 回退一步
                
                paragraph_text = ' '.join(paragraph_lines)
                if paragraph_text:
                    paragraph = MarkdownElement(
                        element_type='paragraph',
                        content=paragraph_text,
                        attributes={}
                    )
                    document.children.append(paragraph)
            
            # 处理代码块
            elif line.startswith('```'):
                language = line[3:].strip()
                code_lines = []
                i += 1
                while i < len(lines) and not lines[i].strip().startswith('```'):
                    code_lines.append(lines[i])
                    i += 1
                
                code_text = '\n'.join(code_lines).strip()
                if code_text:
                    code_block = MarkdownElement(
                        element_type='code_block',
                        content=code_text,
                        attributes={'language': language} if language else {}
                    )
                    document.children.append(code_block)
            
            # 处理引用
            elif line.startswith('>'):
                quote_text = line[1:].strip()
                quote = MarkdownElement(
                    element_type='quote',
                    content=quote_text,
                    attributes={}
                )
                document.children.append(quote)
            
            # 处理表格
            elif '|' in line:
                table_lines = [line]
                i += 1
                # 跳过分隔行
                if i < len(lines) and '|' in lines[i] and re.match(r'^\|?\s*:?-+:?\s*\|', lines[i]):
                    i += 1
                # 收集表格行
                while i < len(lines) and '|' in lines[i]:
                    table_lines.append(lines[i].strip())
                    i += 1
                i -= 1
                
                # 解析表格数据
                table_data = []
                for table_line in table_lines:
                    if '|' in table_line and not re.match(r'^\|?\s*:?-+:?\s*\|', table_line):
                        cells = [cell.strip() for cell in table_line.split('|') if cell.strip()]
                        if cells:
                            table_data.append(cells)
                
                if table_data:
                    table = MarkdownElement(
                        element_type='table',
                        content='',
                        attributes={
                            'rows': len(table_data),
                            'cols': max(len(row) for row in table_data) if table_data else 0,
                            'data': table_data
                        }
                    )
                    document.children.append(table)
            
            i += 1
        
        return document
    
    def extract_metadata(self, text: str) -> Dict[str, Any]:
        """提取文档元数据"""
        metadata = {
            'title': self._extract_title(text),
            'headings': self._extract_headings(text),
            'images': self._extract_images(text),
            'links': self._extract_links(text),
            'code_blocks': self._extract_code_blocks(text),
            'tables': self._extract_tables(text)
        }
        
        return metadata
    
    def _extract_title(self, text: str) -> Optional[str]:
        """提取文档标题"""
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            if line.startswith('# '):
                return line[2:].strip()
        return None
    
    def _extract_headings(self, text: str) -> List[Dict[str, Any]]:
        """提取所有标题"""
        headings = []
        pattern = r'^(#{1,6})\s+(.+)$'
        
        for line_num, line in enumerate(text.split('\n'), 1):
            match = re.match(pattern, line.strip())
            if match:
                level = len(match.group(1))
                title = match.group(2).strip()
                headings.append({
                    'level': level,
                    'title': title,
                    'line': line_num
                })
        
        return headings
    
    def _extract_images(self, text: str) -> List[Dict[str, str]]:
        """提取图片信息"""
        images = []
        pattern = r'!\[([^\]]*)\]\(([^)]+)\)'
        
        for match in re.finditer(pattern, text):
            images.append({
                'alt': match.group(1),
                'src': match.group(2)
            })
        
        return images
    
    def _extract_links(self, text: str) -> List[Dict[str, str]]:
        """提取链接信息"""
        links = []
        pattern = r'\[([^\]]+)\]\(([^)]+)\)'
        
        for match in re.finditer(pattern, text):
            links.append({
                'text': match.group(1),
                'url': match.group(2)
            })
        
        return links
    
    def _extract_code_blocks(self, text: str) -> List[Dict[str, str]]:
        """提取代码块信息"""
        code_blocks = []
        pattern = r'```(\w*)\n([^`]+)```'
        
        for match in re.finditer(pattern, text, re.MULTILINE | re.DOTALL):
            code_blocks.append({
                'language': match.group(1) or 'text',
                'code': match.group(2).strip()
            })
        
        return code_blocks
    
    def _extract_tables(self, text: str) -> List[Dict[str, Any]]:
        """提取表格信息"""
        tables = []
        lines = text.split('\n')
        
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if '|' in line and i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                if re.match(r'^\|?\s*:?-+:?\s*\|', next_line):
                    # 找到表格
                    table_lines = [line]
                    i += 1
                    
                    # 跳过分隔行
                    i += 1
                    
                    # 收集表格行
                    while i < len(lines) and '|' in lines[i]:
                        table_lines.append(lines[i].strip())
                        i += 1
                    
                    tables.append({
                        'rows': len(table_lines),
                        'content': '\n'.join(table_lines)
                    })
                    continue
            i += 1
        
        return tables