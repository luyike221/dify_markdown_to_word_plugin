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
        # 标准化换行符
        text = text.replace('\r\n', '\n').replace('\r', '\n')
        
        # 处理数学公式
        text = self._process_math_formulas(text)
        
        # 处理图片链接
        text = self._process_images(text)
        
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
    
    def _build_document_tree(self, html_content: str, original_text: str) -> MarkdownElement:
        """构建文档树"""
        # 这里需要进一步解析HTML内容，构建结构化的文档树
        # 暂时返回基本结构
        document = MarkdownElement(
            element_type='document',
            content=html_content,
            attributes={'original_text': original_text}
        )
        
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