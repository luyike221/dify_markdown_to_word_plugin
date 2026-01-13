#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置数据模型
定义所有样式配置的数据结构
"""

from dataclasses import dataclass, field
from typing import Optional, Dict, List, Any
import re


@dataclass
class FontStyle:
    """字体样式"""
    family: str = "微软雅黑"
    size: int = 12                    # pt
    color: str = "#000000"
    bold: bool = False
    italic: bool = False
    underline: bool = False


@dataclass
class ParagraphStyle:
    """段落样式"""
    alignment: str = "left"           # left, center, right, justify
    line_spacing: float = 1.15        # >=20: 固定值(磅), <20: 倍数
    space_before: float = 0           # pt
    space_after: float = 0            # pt
    left_indent: float = 0            # cm
    right_indent: float = 0           # cm
    first_line_indent: float = 0      # cm, 0表示根据字体大小动态计算
    keep_together: bool = False
    keep_with_next: bool = False
    page_break_before: bool = False


@dataclass
class ElementStyle:
    """元素完整样式"""
    font: FontStyle = field(default_factory=FontStyle)
    paragraph: ParagraphStyle = field(default_factory=ParagraphStyle)
    background_color: Optional[str] = None
    
    def merge(self, other: 'ElementStyle') -> 'ElementStyle':
        """合并样式（other 覆盖 self）
        
        Args:
            other: 要合并的样式
            
        Returns:
            合并后的新样式
        """
        import copy
        result = copy.deepcopy(self)
        
        # 合并字体样式（只覆盖非默认值）
        if other.font:
            default_font = FontStyle()
            for key, value in other.font.__dict__.items():
                other_value = getattr(other.font, key)
                default_value = getattr(default_font, key)
                # 如果 other 的值不是默认值，则覆盖
                if other_value != default_value:
                    setattr(result.font, key, other_value)
        
        # 合并段落样式
        if other.paragraph:
            default_para = ParagraphStyle()
            for key, value in other.paragraph.__dict__.items():
                other_value = getattr(other.paragraph, key)
                default_value = getattr(default_para, key)
                if other_value != default_value:
                    setattr(result.paragraph, key, other_value)
        
        # 合并背景色
        if other.background_color:
            result.background_color = other.background_color
        
        return result


@dataclass
class PageStyle:
    """页面样式"""
    size: str = "A4"                  # A4, Letter, Legal
    width: float = 21.0               # cm
    height: float = 29.7              # cm
    orientation: str = "portrait"     # portrait, landscape
    margin_top: float = 2.5           # cm
    margin_bottom: float = 2.5        # cm
    margin_left: float = 3.0          # cm
    margin_right: float = 2.5         # cm


@dataclass
class HeadingStyles:
    """标题样式集合（支持继承）"""
    default: ElementStyle = field(default_factory=ElementStyle)
    h1: Optional[ElementStyle] = None
    h2: Optional[ElementStyle] = None
    h3: Optional[ElementStyle] = None
    h4: Optional[ElementStyle] = None
    h5: Optional[ElementStyle] = None
    h6: Optional[ElementStyle] = None
    
    def get(self, level: int) -> ElementStyle:
        """获取指定级别的标题样式（带继承）
        
        Args:
            level: 标题级别 (1-6)
        
        Returns:
            合并后的完整样式
        """
        # 获取对应级别的样式
        level_style = getattr(self, f'h{level}', None)
        
        # 如果没有定义，直接返回默认样式
        if level_style is None:
            return self.default
        
        # 合并默认样式和级别样式
        return self.default.merge(level_style)


@dataclass
class TableStyle:
    """表格样式"""
    border_width: float = 1.0         # pt
    border_color: str = "#424242"
    cell_padding: int = 6             # pt
    cell_alignment: str = "center"
    cell_font_family: str = "微软雅黑"
    cell_font_size: int = 9
    header_background: str = "#f44336"
    header_font_color: str = "#ffffff"
    header_font_bold: bool = True
    alternate_row_color: str = "#fff3e0"


@dataclass
class ChartStyle:
    """图表样式"""
    width: float = 14.0               # cm (生成时)
    insert_width: float = 14.0        # cm (插入Word时)
    dpi: int = 150
    background_color: str = "#FFFFFF"
    colors: List[str] = field(default_factory=list)
    font_sizes: Dict[str, int] = field(default_factory=dict)
    add_title: bool = False
    pie_threshold: float = 8.0        # 饼图标注阈值（百分比），小于此值的切片数字标识会移到外部并使用引线
    
    def __post_init__(self):
        """初始化默认值"""
        if not self.colors:
            self.colors = [
                "#2E86AB", "#A23B72", "#F18F01", "#C73E1D",
                "#6A994E", "#BC4749", "#F77F00", "#FCBF49",
                "#06A77D", "#7209B7", "#3A86FF", "#FF006E"
            ]
        if not self.font_sizes:
            self.font_sizes = {
                "title": 14,
                "label": 10,
                "legend": 10,
                "value": 9,
                "y_axis": 12
            }


@dataclass
class StyleConfig:
    """完整样式配置"""
    # 页面
    page: PageStyle = field(default_factory=PageStyle)
    
    # 正文
    body: ElementStyle = field(default_factory=ElementStyle)
    
    # 标题
    headings: HeadingStyles = field(default_factory=HeadingStyles)
    
    # 代码
    code_inline: ElementStyle = field(default_factory=ElementStyle)
    code_block: ElementStyle = field(default_factory=ElementStyle)
    
    # 表格
    table: TableStyle = field(default_factory=TableStyle)
    
    # 引用
    quote: ElementStyle = field(default_factory=ElementStyle)
    
    # 列表
    list_indent: float = 1.27         # cm
    
    # 图表
    chart: ChartStyle = field(default_factory=ChartStyle)
    
    # 页码
    enable_page_numbers: bool = True
    
    def __post_init__(self):
        """初始化默认样式"""
        # 设置正文默认样式
        if not hasattr(self.body.font, 'family') or self.body.font.family == "微软雅黑":
            self.body.font = FontStyle(
                family="宋体",
                size=14,
                color="#000000"
            )
            self.body.paragraph = ParagraphStyle(
                line_spacing=28,  # 28磅固定值
                space_before=0,
                space_after=0,
                first_line_indent=0  # 动态计算两个字符宽度
            )
        
        # 设置标题默认样式
        if not self.headings.default.font.family or self.headings.default.font.family == "微软雅黑":
            # 通用标题默认样式
            self.headings.default = ElementStyle(
                font=FontStyle(
                    family="黑体",
                    size=15,
                    bold=True,
                    color="#000000"
                ),
                paragraph=ParagraphStyle(
                    line_spacing=28,
                    space_before=0,
                    space_after=0,
                    keep_with_next=True
                )
            )
            
            # H1 特殊样式
            self.headings.h1 = ElementStyle(
                font=FontStyle(
                    family="宋体",
                    size=22,
                    bold=False,
                    color="#000000"
                ),
                paragraph=ParagraphStyle(
                    alignment="center",
                    line_spacing=1.25,
                    space_before=0,
                    space_after=0,
                    keep_with_next=True
                )
            )
            
            # H2 样式
            self.headings.h2 = ElementStyle(
                font=FontStyle(
                    family="黑体",
                    size=16,
                    bold=True,
                    color="#000000"
                ),
                paragraph=ParagraphStyle(
                    line_spacing=28,
                    space_before=0,
                    space_after=0,
                    keep_with_next=True
                )
            )
            
            # H3 样式
            self.headings.h3 = ElementStyle(
                font=FontStyle(
                    family="黑体",
                    size=15,
                    bold=True,
                    color="#000000"
                ),
                paragraph=ParagraphStyle(
                    line_spacing=28,
                    space_before=0,
                    space_after=0,
                    keep_with_next=True
                )
            )
        
        # 设置代码样式
        if self.code_inline.font.family == "微软雅黑":
            self.code_inline.font = FontStyle(
                family="Consolas",
                size=10,
                color="#d32f2f"
            )
            self.code_inline.background_color = "#f5f5f5"
        
        if self.code_block.font.family == "微软雅黑":
            self.code_block.font = FontStyle(
                family="Consolas",
                size=9,
                color="#333333"
            )
            self.code_block.paragraph = ParagraphStyle(
                line_spacing=1.2,
                left_indent=0.2
            )
            self.code_block.background_color = "#f5f5f5"
        
        # 设置引用样式
        if self.quote.font.family == "微软雅黑":
            self.quote.font = FontStyle(
                family="宋体",
                size=10,
                color="#666666",
                italic=True
            )
            self.quote.paragraph = ParagraphStyle(
                left_indent=0.5,
                line_spacing=1.5
            )
            self.quote.background_color = "#fff3e0"

