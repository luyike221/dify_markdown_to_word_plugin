#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置管理器
负责加载、解析、合并和管理样式配置
"""

from pathlib import Path
from typing import Optional, Dict, Any
import yaml
import json

from .models import (
    StyleConfig,
    ElementStyle,
    FontStyle,
    ParagraphStyle,
    PageStyle,
    HeadingStyles,
    ChartStyle,
    TableStyle,
)


class ConfigManager:
    """统一配置管理器"""
    
    def __init__(self, config_dir: Optional[Path] = None):
        """初始化配置管理器
        
        Args:
            config_dir: 配置文件目录，默认为项目根目录的 config/
        """
        if config_dir is None:
            # 默认配置目录
            current_file = Path(__file__)
            project_root = current_file.parent.parent.parent
            config_dir = project_root / "config"
        
        self.config_dir = Path(config_dir)
        self.config_dir.mkdir(exist_ok=True, parents=True)
        
        # 配置文件路径
        self.default_config_file = self.config_dir / "style.yaml"
        self.themes_dir = self.config_dir / "themes"
        self.themes_dir.mkdir(exist_ok=True, parents=True)
    
    def load_config(
        self,
        json_config: Optional[str] = None,
        theme: Optional[str] = None
    ) -> StyleConfig:
        """加载配置
        
        优先级（从低到高）：
        1. 内置默认配置（StyleConfig 的 __post_init__）
        2. 系统配置文件 (config/style.yaml)
        3. 主题配置 (config/themes/{theme}.yaml)
        4. JSON 配置 (API 传入的 style_config 参数)
        
        Args:
            json_config: JSON 格式的配置字符串
            theme: 主题名称
        
        Returns:
            合并后的完整配置
        """
        # 1. 从默认配置开始
        config = StyleConfig()
        
        # 2. 加载系统配置文件（如果存在）
        if self.default_config_file.exists():
            system_config_dict = self._load_yaml(self.default_config_file)
            config = self._apply_dict_to_config(config, system_config_dict)
        
        # 3. 加载主题配置（如果指定）
        if theme and theme != 'default':
            theme_dict = self._load_theme(theme)
            if theme_dict:
                config = self._apply_dict_to_config(config, theme_dict)
        
        # 4. 应用 JSON 配置（最高优先级）
        if json_config:
            json_dict = self._parse_json_config(json_config)
            config = self._apply_dict_to_config(config, json_dict)
        
        return config
    
    def _parse_json_config(self, json_str: str) -> Dict[str, Any]:
        """解析 JSON 配置字符串
        
        支持的格式：
        1. 旧格式：{"text_style": {...}, "graph_style": {...}}
        2. 新格式：{"body": {...}, "headings": {...}, ...}
        
        Args:
            json_str: JSON 字符串
        
        Returns:
            标准化的配置字典
        """
        try:
            # 清理 JSON 字符串
            if isinstance(json_str, str):
                json_str = json_str.replace('\\n', '\n').replace('\\t', '\t')
                json_str = json_str.replace('\xa0', ' ').strip()
            
            # 解析 JSON
            data = json.loads(json_str) if isinstance(json_str, str) else json_str
            
            # 检测格式并转换
            if "text_style" in data or "graph_style" in data:
                # 旧格式，需要转换
                return self._migrate_old_format(data)
            else:
                # 新格式，直接返回
                return data
        
        except json.JSONDecodeError as e:
            print(f"JSON 解析失败: {e}")
            return {}
        except Exception as e:
            print(f"配置解析错误: {e}")
            return {}
    
    def _migrate_old_format(self, old_config: Dict[str, Any]) -> Dict[str, Any]:
        """将旧格式配置迁移到新格式
        
        旧格式：
        {
          "text_style": {
            "font_family": "宋体",
            "font_size": 12,
            "heading1_font_family": "方正小标宋简体",
            ...
          },
          "graph_style": {...}
        }
        
        新格式：
        {
          "body": {"font": {...}, "paragraph": {...}},
          "headings": {"h1": {...}, "h2": {...}},
          "chart": {...}
        }
        
        Args:
            old_config: 旧格式配置
        
        Returns:
            新格式配置
        """
        new_config = {}
        
        # 处理 text_style
        if "text_style" in old_config:
            text_style = old_config["text_style"]
            
            # 正文样式
            new_config["body"] = {
                "font": {
                    "family": text_style.get("font_family", "宋体"),
                    "size": int(text_style.get("font_size", 12)),
                    "color": text_style.get("font_color", "#000000"),
                },
                "paragraph": {
                    "line_spacing": float(text_style.get("line_spacing", 1.5)),
                }
            }
            
            # 页面设置
            margins = text_style.get("page_margins", 2.5)
            new_config["page"] = {
                "size": text_style.get("paper_size", "A4"),
                "margin_top": margins,
                "margin_bottom": margins,
                "margin_left": margins,
                "margin_right": margins,
            }
            
            # 标题样式
            new_config["headings"] = {}
            
            # 处理 H1-H6
            for level in range(1, 7):
                # 确定配置前缀
                if level <= 3:
                    prefix = f"heading{level}_"
                else:
                    prefix = "heading_"
                
                heading_data = {}
                
                # 字体配置
                font_family = text_style.get(f"{prefix}font_family")
                font_size = text_style.get(f"{prefix}font_size")
                font_color = text_style.get(f"{prefix}font_color")
                font_bold = text_style.get(f"{prefix}bold")
                
                if font_family or font_size or font_color or font_bold is not None:
                    heading_data["font"] = {}
                    if font_family:
                        heading_data["font"]["family"] = font_family
                    if font_size:
                        heading_data["font"]["size"] = int(font_size)
                    if font_color:
                        heading_data["font"]["color"] = font_color
                    if font_bold is not None:
                        heading_data["font"]["bold"] = bool(font_bold)
                
                # 段落配置
                line_spacing = text_style.get(f"{prefix}line_spacing")
                if line_spacing:
                    heading_data["paragraph"] = {
                        "line_spacing": float(line_spacing)
                    }
                
                if heading_data:
                    new_config["headings"][f"h{level}"] = heading_data
        
        # 处理 graph_style
        if "graph_style" in old_config:
            graph_style = old_config["graph_style"]
            new_config["chart"] = {
                "width": float(graph_style.get("chart_width", 14.0)),
                "insert_width": float(graph_style.get("chart_insert_width", 14.0)),
                "dpi": int(graph_style.get("chart_dpi", 150)),
                "background_color": graph_style.get("background_color", "#FFFFFF"),
                "colors": graph_style.get("colors", []),
                "font_sizes": graph_style.get("font_sizes", {}),
                "add_title": bool(graph_style.get("add_chart_title", False)),
            }
        
        return new_config
    
    def _apply_dict_to_config(
        self,
        config: StyleConfig,
        config_dict: Dict[str, Any]
    ) -> StyleConfig:
        """将字典配置应用到 StyleConfig 对象
        
        Args:
            config: 基础配置对象
            config_dict: 要应用的配置字典
        
        Returns:
            更新后的配置对象
        """
        # 应用页面配置
        if "page" in config_dict:
            page_data = config_dict["page"]
            if isinstance(page_data, dict):
                for key, value in page_data.items():
                    if hasattr(config.page, key):
                        setattr(config.page, key, value)
        
        # 应用正文配置
        if "body" in config_dict:
            config.body = self._dict_to_element_style(
                config_dict["body"],
                config.body
            )
        
        # 应用标题配置
        if "headings" in config_dict:
            headings_data = config_dict["headings"]
            
            # 默认标题样式
            if "default" in headings_data:
                config.headings.default = self._dict_to_element_style(
                    headings_data["default"],
                    config.headings.default
                )
            
            # 各级标题
            for level in range(1, 7):
                h_key = f"h{level}"
                if h_key in headings_data:
                    current_style = getattr(config.headings, h_key, None)
                    if current_style is None:
                        current_style = ElementStyle()
                    
                    new_style = self._dict_to_element_style(
                        headings_data[h_key],
                        current_style
                    )
                    setattr(config.headings, h_key, new_style)
        
        # 应用代码配置
        if "code_inline" in config_dict:
            config.code_inline = self._dict_to_element_style(
                config_dict["code_inline"],
                config.code_inline
            )
        
        if "code_block" in config_dict:
            config.code_block = self._dict_to_element_style(
                config_dict["code_block"],
                config.code_block
            )
        
        # 应用表格配置
        if "table" in config_dict:
            table_data = config_dict["table"]
            if isinstance(table_data, dict):
                for key, value in table_data.items():
                    if hasattr(config.table, key):
                        setattr(config.table, key, value)
        
        # 应用引用配置
        if "quote" in config_dict:
            config.quote = self._dict_to_element_style(
                config_dict["quote"],
                config.quote
            )
        
        # 应用图表配置
        if "chart" in config_dict:
            chart_data = config_dict["chart"]
            if isinstance(chart_data, dict):
                for key, value in chart_data.items():
                    if hasattr(config.chart, key):
                        setattr(config.chart, key, value)
        
        # 应用列表缩进
        if "list_indent" in config_dict:
            config.list_indent = float(config_dict["list_indent"])
        
        # 应用页码设置
        if "enable_page_numbers" in config_dict:
            config.enable_page_numbers = bool(config_dict["enable_page_numbers"])
        
        return config
    
    def _dict_to_element_style(
        self,
        data: Dict[str, Any],
        base_style: Optional[ElementStyle] = None
    ) -> ElementStyle:
        """将字典转换为 ElementStyle
        
        Args:
            data: 样式字典
            base_style: 基础样式（用于继承）
        
        Returns:
            ElementStyle 对象
        """
        if base_style is None:
            base_style = ElementStyle()
        
        # 应用字体样式
        if "font" in data:
            font_data = data["font"]
            if isinstance(font_data, dict):
                for key, value in font_data.items():
                    if hasattr(base_style.font, key):
                        setattr(base_style.font, key, value)
        
        # 应用段落样式
        if "paragraph" in data:
            para_data = data["paragraph"]
            if isinstance(para_data, dict):
                for key, value in para_data.items():
                    if hasattr(base_style.paragraph, key):
                        setattr(base_style.paragraph, key, value)
        
        # 应用背景色
        if "background_color" in data:
            base_style.background_color = data["background_color"]
        
        return base_style
    
    def _load_yaml(self, file_path: Path) -> Dict[str, Any]:
        """加载 YAML 文件"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f) or {}
        except Exception as e:
            print(f"加载 YAML 文件失败 {file_path}: {e}")
            return {}
    
    def _load_theme(self, theme_name: str) -> Optional[Dict[str, Any]]:
        """加载主题配置"""
        theme_file = self.themes_dir / f"{theme_name}.yaml"
        if theme_file.exists():
            return self._load_yaml(theme_file)
        return None
    
    def save_config(
        self,
        config: StyleConfig,
        file_path: Optional[Path] = None
    ):
        """保存配置到 YAML 文件
        
        Args:
            config: 配置对象
            file_path: 保存路径，默认为 config/style.yaml
        """
        if file_path is None:
            file_path = self.default_config_file
        
        # 转换为字典
        from dataclasses import asdict
        config_dict = asdict(config)
        
        # 保存为 YAML
        with open(file_path, 'w', encoding='utf-8') as f:
            yaml.dump(
                config_dict,
                f,
                allow_unicode=True,
                indent=2,
                default_flow_style=False
            )

