#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
样式处理器模块
负责处理Word文档的样式配置、主题管理和格式化
"""

import json
import os
from typing import Dict, List, Any, Optional, Union
from dataclasses import dataclass, asdict
from pathlib import Path

try:
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.style import WD_STYLE_TYPE
except ImportError:
    raise ImportError("请安装python-docx库: pip install python-docx")


@dataclass
class FontStyle:
    """字体样式配置"""
    name: str = "微软雅黑"
    size: int = 12
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: Optional[str] = None
    highlight: Optional[str] = None


@dataclass
class ParagraphStyle:
    """段落样式配置"""
    alignment: str = "left"  # left, center, right, justify
    line_spacing: float = 1.15
    space_before: int = 0
    space_after: int = 6
    left_indent: float = 0.0
    right_indent: float = 0.0
    first_line_indent: float = 0.0
    keep_together: bool = False
    keep_with_next: bool = False
    page_break_before: bool = False


@dataclass
class BorderStyle:
    """边框样式配置"""
    top: Optional[Dict[str, Any]] = None
    bottom: Optional[Dict[str, Any]] = None
    left: Optional[Dict[str, Any]] = None
    right: Optional[Dict[str, Any]] = None
    
    def __post_init__(self):
        default_border = {
            'style': 'single',
            'width': 1,
            'color': '#000000'
        }
        
        if self.top is None:
            self.top = {}
        if self.bottom is None:
            self.bottom = {}
        if self.left is None:
            self.left = {}
        if self.right is None:
            self.right = {}


@dataclass
class ElementStyle:
    """元素样式配置"""
    font: FontStyle
    paragraph: ParagraphStyle
    border: Optional[BorderStyle] = None
    background_color: Optional[str] = None
    
    def __post_init__(self):
        if isinstance(self.font, dict):
            self.font = FontStyle(**self.font)
        if isinstance(self.paragraph, dict):
            self.paragraph = ParagraphStyle(**self.paragraph)
        if self.border and isinstance(self.border, dict):
            self.border = BorderStyle(**self.border)


class StyleTheme:
    """样式主题"""
    
    def __init__(self, name: str, styles: Dict[str, ElementStyle]):
        self.name = name
        self.styles = styles
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'StyleTheme':
        """从字典创建主题"""
        name = data.get('name', 'default')
        styles = {}
        
        for style_name, style_data in data.get('styles', {}).items():
            styles[style_name] = ElementStyle(**style_data)
        
        return cls(name, styles)
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            'name': self.name,
            'styles': {name: asdict(style) for name, style in self.styles.items()}
        }
    
    def get_style(self, element_type: str) -> Optional[ElementStyle]:
        """获取指定元素的样式"""
        return self.styles.get(element_type)


class StyleHandler:
    """样式处理器"""
    
    def __init__(self, config_path: Optional[str] = None):
        """初始化样式处理器
        
        Args:
            config_path: 样式配置文件路径
        """
        self.config_path = config_path
        self.themes = {}
        self.current_theme = None
        
        # 加载默认主题
        self._load_default_themes()
        
        # 加载自定义配置
        if config_path and os.path.exists(config_path):
            self.load_config(config_path)
        
        # 设置默认主题
        self.set_theme('default')
    
    def _load_default_themes(self):
        """加载默认主题"""
        # 默认主题（使用告警报告样式）
        default_styles = {
            'normal': ElementStyle(
                font=FontStyle(name="宋体", size=14),  # 四号字体
                paragraph=ParagraphStyle(
                    line_spacing=28,  # 28磅固定值行距
                    space_before=0,
                    space_after=0,
                    first_line_indent=0  # 0表示根据字体大小动态计算两个字符的宽度
                )
            ),
            'heading1': ElementStyle(
                font=FontStyle(name="宋体", size=22, bold=False, color="#000000"),  # 二号字体，黑色
                paragraph=ParagraphStyle(
                    alignment="center",  # 居中
                    line_spacing=1.25,  # 1.25倍行距
                    space_before=0,
                    space_after=0,
                    keep_with_next=True
                )
            ),
            'heading2': ElementStyle(
                font=FontStyle(name="黑体", size=16, bold=True, color="#000000"),  # 三号字体，黑色
                paragraph=ParagraphStyle(
                    line_spacing=28,  # 28磅固定值行距
                    space_before=0,
                    space_after=0,
                    keep_with_next=True
                )
            ),
            'heading3': ElementStyle(
                font=FontStyle(name="黑体", size=15, bold=True, color="#000000"),  # 小三号字体，黑色
                paragraph=ParagraphStyle(
                    line_spacing=28,  # 28磅固定值行距
                    space_before=0,
                    space_after=0,
                    keep_with_next=True
                )
            ),
            'heading4': ElementStyle(
                font=FontStyle(name="黑体", size=15, bold=True, color="#000000"),  # 小三号字体，黑色
                paragraph=ParagraphStyle(
                    line_spacing=28,  # 28磅固定值行距
                    space_before=0,
                    space_after=0,
                    keep_with_next=True
                )
            ),
            'heading5': ElementStyle(
                font=FontStyle(name="黑体", size=15, bold=True, color="#000000"),  # 小三号字体，黑色
                paragraph=ParagraphStyle(
                    line_spacing=28,  # 28磅固定值行距
                    space_before=0,
                    space_after=0,
                    keep_with_next=True
                )
            ),
            'heading6': ElementStyle(
                font=FontStyle(name="黑体", size=15, bold=True, color="#000000"),  # 小三号字体，黑色
                paragraph=ParagraphStyle(
                    line_spacing=28,  # 28磅固定值行距
                    space_before=0,
                    space_after=0,
                    keep_with_next=True
                )
            ),
            'code_inline': ElementStyle(
                font=FontStyle(name="Consolas", size=10, color="#333333"),
                paragraph=ParagraphStyle(),
                background_color="#f8f8f8"
            ),
            'code_block': ElementStyle(
                font=FontStyle(name="Consolas", size=10, color="#333333"),
                paragraph=ParagraphStyle(left_indent=0.2, space_before=6, space_after=6),
                background_color="#f8f8f8",
                border=BorderStyle(left={'style': 'single', 'width': 2, 'color': '#cccccc'})
            ),
            'quote': ElementStyle(
                font=FontStyle(name="微软雅黑", size=12, italic=True, color="#666666"),
                paragraph=ParagraphStyle(left_indent=0.5, space_before=6, space_after=6),
                border=BorderStyle(left={'style': 'single', 'width': 3, 'color': '#cccccc'})
            ),
            'table_header': ElementStyle(
                font=FontStyle(name="微软雅黑", size=9, bold=True),  # 小五字体
                paragraph=ParagraphStyle(alignment="center"),
                background_color="#f44336"
            ),
            'table_cell': ElementStyle(
                font=FontStyle(name="微软雅黑", size=9),  # 小五字体
                paragraph=ParagraphStyle(alignment="center")
            ),
            'caption': ElementStyle(
                font=FontStyle(name="微软雅黑", size=10, italic=True, color="#666666"),
                paragraph=ParagraphStyle(alignment="center", space_before=3, space_after=6)
            ),
            'list_bullet': ElementStyle(
                font=FontStyle(name="微软雅黑", size=12),
                paragraph=ParagraphStyle(left_indent=0.25, space_after=3)
            ),
            'list_number': ElementStyle(
                font=FontStyle(name="微软雅黑", size=12),
                paragraph=ParagraphStyle(left_indent=0.25, space_after=3)
            )
        }
        
        self.themes['default'] = StyleTheme('default', default_styles)
        
        # 学术论文主题
        academic_styles = default_styles.copy()
        academic_styles.update({
            'normal': ElementStyle(
                font=FontStyle(name="Times New Roman", size=12),
                paragraph=ParagraphStyle(line_spacing=2.0, space_after=0, alignment="justify")
            ),
            'heading1': ElementStyle(
                font=FontStyle(name="Times New Roman", size=16, bold=True),
                paragraph=ParagraphStyle(alignment="center", space_before=24, space_after=12, keep_with_next=True)
            ),
            'heading2': ElementStyle(
                font=FontStyle(name="Times New Roman", size=14, bold=True),
                paragraph=ParagraphStyle(space_before=18, space_after=6, keep_with_next=True)
            )
        })
        
        self.themes['academic'] = StyleTheme('academic', academic_styles)
        
        # 商务报告主题
        business_styles = default_styles.copy()
        business_styles.update({
            'normal': ElementStyle(
                font=FontStyle(name="Calibri", size=11),
                paragraph=ParagraphStyle(line_spacing=1.15, space_after=6)
            ),
            'heading1': ElementStyle(
                font=FontStyle(name="Calibri", size=18, bold=True, color="#2F5597"),
                paragraph=ParagraphStyle(space_before=12, space_after=6, keep_with_next=True)
            ),
            'heading2': ElementStyle(
                font=FontStyle(name="Calibri", size=14, bold=True, color="#2F5597"),
                paragraph=ParagraphStyle(space_before=10, space_after=6, keep_with_next=True)
            )
        })
        
        self.themes['business'] = StyleTheme('business', business_styles)
    
    def load_config(self, config_path: str):
        """加载样式配置文件"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            
            # 加载主题
            if 'themes' in config_data:
                for theme_data in config_data['themes']:
                    theme = StyleTheme.from_dict(theme_data)
                    self.themes[theme.name] = theme
            
        except Exception as e:
            print(f"加载样式配置失败: {e}")
    
    def save_config(self, config_path: str):
        """保存样式配置文件"""
        try:
            config_data = {
                'themes': [theme.to_dict() for theme in self.themes.values()]
            }
            
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            print(f"保存样式配置失败: {e}")
    
    def set_theme(self, theme_name: str) -> bool:
        """设置当前主题"""
        if theme_name in self.themes:
            self.current_theme = self.themes[theme_name]
            return True
        return False
    
    def get_current_theme(self) -> Optional[StyleTheme]:
        """获取当前主题"""
        return self.current_theme
    
    def get_style(self, element_type: str) -> Optional[ElementStyle]:
        """获取指定元素的样式"""
        if self.current_theme:
            return self.current_theme.get_style(element_type)
        return None
    
    def add_theme(self, theme: StyleTheme):
        """添加主题"""
        self.themes[theme.name] = theme
    
    def remove_theme(self, theme_name: str) -> bool:
        """删除主题"""
        if theme_name in self.themes and theme_name != 'default':
            del self.themes[theme_name]
            if self.current_theme and self.current_theme.name == theme_name:
                self.set_theme('default')
            return True
        return False
    
    def list_themes(self) -> List[str]:
        """列出所有主题名称"""
        return list(self.themes.keys())
    
    def create_custom_style(self, name: str, font_config: Dict[str, Any], 
                           paragraph_config: Dict[str, Any], 
                           border_config: Optional[Dict[str, Any]] = None,
                           background_color: Optional[str] = None) -> ElementStyle:
        """创建自定义样式"""
        font_style = FontStyle(**font_config)
        paragraph_style = ParagraphStyle(**paragraph_config)
        border_style = BorderStyle(**border_config) if border_config else None
        
        return ElementStyle(
            font=font_style,
            paragraph=paragraph_style,
            border=border_style,
            background_color=background_color
        )
    
    def apply_style_to_paragraph(self, paragraph, element_type: str):
        """将样式应用到段落"""
        style = self.get_style(element_type)
        if not style:
            return
        
        # 应用字体样式
        self._apply_font_style(paragraph, style.font)
        
        # 应用段落样式，传递字体大小用于动态计算首行缩进
        self._apply_paragraph_style(paragraph, style.paragraph, font_size=style.font.size)
        
        # 应用边框样式
        if style.border:
            self._apply_border_style(paragraph, style.border)
        
        # 应用背景色
        if style.background_color:
            self._apply_background_color(paragraph, style.background_color)
    
    def _apply_font_style(self, paragraph, font_style: FontStyle):
        """应用字体样式"""
        for run in paragraph.runs:
            font = run.font
            font.name = font_style.name
            font.size = Pt(font_style.size)
            font.bold = font_style.bold
            font.italic = font_style.italic
            font.underline = font_style.underline
            
            if font_style.color:
                color_hex = font_style.color.replace('#', '')
                if len(color_hex) == 6:
                    r = int(color_hex[0:2], 16)
                    g = int(color_hex[2:4], 16)
                    b = int(color_hex[4:6], 16)
                    font.color.rgb = RGBColor(r, g, b)
    
    def _apply_paragraph_style(self, paragraph, paragraph_style: ParagraphStyle, font_size: Optional[int] = None):
        """应用段落样式
        
        Args:
            paragraph: Word段落对象
            paragraph_style: 段落样式配置
            font_size: 字体大小（磅），用于动态计算首行缩进
        """
        from docx.enum.text import WD_LINE_SPACING
        
        pf = paragraph.paragraph_format
        
        # 对齐方式
        alignment_map = {
            'left': WD_ALIGN_PARAGRAPH.LEFT,
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT,
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        paragraph.alignment = alignment_map.get(paragraph_style.alignment, WD_ALIGN_PARAGRAPH.LEFT)
        
        # 行距处理
        # 如果 line_spacing >= 20，认为是固定值（磅），否则认为是倍数
        if paragraph_style.line_spacing >= 20:
            # 固定值行距（磅）
            pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            pf.line_spacing = Pt(paragraph_style.line_spacing)
        else:
            # 倍数行距
            pf.line_spacing = paragraph_style.line_spacing
        
        # 段前段后间距
        pf.space_before = Pt(paragraph_style.space_before)
        pf.space_after = Pt(paragraph_style.space_after)
        
        # 缩进
        # 注意：配置中的值单位是厘米，需要转换为英寸
        # 1厘米 = 1/2.54 英寸
        pf.left_indent = Inches(paragraph_style.left_indent / 2.54)
        pf.right_indent = Inches(paragraph_style.right_indent / 2.54)
        
        # 首行缩进：如果配置值为0且提供了字体大小，则动态计算两个字符的宽度
        if paragraph_style.first_line_indent == 0 and font_size:
            # 两个字符宽度 = 字体大小 × 2（磅），转换为英寸
            # 1磅 = 1/72 英寸
            first_line_indent_inches = (font_size * 2) / 72.0
            pf.first_line_indent = Inches(first_line_indent_inches)
        else:
            # 使用配置值（厘米转英寸）
            pf.first_line_indent = Inches(paragraph_style.first_line_indent / 2.54)
        
        # 分页控制
        pf.keep_together = paragraph_style.keep_together
        pf.keep_with_next = paragraph_style.keep_with_next
        pf.page_break_before = paragraph_style.page_break_before
    
    def _apply_border_style(self, paragraph, border_style: BorderStyle):
        """应用边框样式"""
        # 这里需要使用python-docx的底层API来设置边框
        # 具体实现较为复杂，暂时简化处理
        pass
    
    def _apply_background_color(self, paragraph, background_color: str):
        """应用背景色"""
        # 这里需要使用python-docx的底层API来设置背景色
        # 具体实现较为复杂，暂时简化处理
        pass
    
    def get_style_preview(self, element_type: str) -> Dict[str, Any]:
        """获取样式预览信息"""
        style = self.get_style(element_type)
        if not style:
            return {}
        
        return {
            'font': asdict(style.font),
            'paragraph': asdict(style.paragraph),
            'border': asdict(style.border) if style.border else None,
            'background_color': style.background_color
        }
    
    def validate_style(self, style: ElementStyle) -> List[str]:
        """验证样式配置"""
        errors = []
        
        # 验证字体大小
        if style.font.size <= 0 or style.font.size > 72:
            errors.append("字体大小应在1-72之间")
        
        # 验证颜色格式
        if style.font.color and not self._is_valid_color(style.font.color):
            errors.append("字体颜色格式无效")
        
        if style.background_color and not self._is_valid_color(style.background_color):
            errors.append("背景色格式无效")
        
        # 验证行距
        if style.paragraph.line_spacing <= 0:
            errors.append("行距必须大于0")
        
        return errors
    
    def _is_valid_color(self, color: str) -> bool:
        """验证颜色格式"""
        if not color.startswith('#'):
            return False
        
        hex_part = color[1:]
        if len(hex_part) not in [3, 6]:
            return False
        
        try:
            int(hex_part, 16)
            return True
        except ValueError:
            return False