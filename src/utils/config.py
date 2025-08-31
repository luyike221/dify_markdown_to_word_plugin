#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置管理模块
负责系统配置、用户配置和转换配置的管理
"""

import os
import json
import yaml
from typing import Dict, Any, Optional, Union
from pathlib import Path
from dataclasses import dataclass, asdict
from enum import Enum


class OutputFormat(Enum):
    """输出格式枚举"""
    DOCX = "docx"
    DOC = "doc"
    PDF = "pdf"
    HTML = "html"


class StyleTheme(Enum):
    """样式主题枚举"""
    DEFAULT = "default"
    ACADEMIC = "academic"
    BUSINESS = "business"
    MINIMAL = "minimal"
    MODERN = "modern"
    CLASSIC = "classic"


class LogLevel(Enum):
    """日志级别枚举"""
    DEBUG = "DEBUG"
    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"
    CRITICAL = "CRITICAL"


@dataclass
class FontConfig:
    """字体配置"""
    name: str = "微软雅黑"
    size: int = 12
    bold: bool = False
    italic: bool = False
    color: str = "#000000"


@dataclass
class PageConfig:
    """页面配置"""
    width: float = 21.0  # cm
    height: float = 29.7  # cm
    margin_top: float = 2.54  # cm
    margin_bottom: float = 2.54  # cm
    margin_left: float = 3.18  # cm
    margin_right: float = 3.18  # cm
    orientation: str = "portrait"  # portrait or landscape


@dataclass
class StyleConfig:
    """样式配置"""
    theme: str = StyleTheme.DEFAULT.value
    font: FontConfig = None
    heading_font: FontConfig = None
    code_font: FontConfig = None
    page: PageConfig = None
    line_spacing: float = 1.15
    paragraph_spacing: float = 6.0
    enable_page_numbers: bool = True
    enable_toc: bool = True
    enable_header: bool = False
    enable_footer: bool = False
    
    def __post_init__(self):
        if self.font is None:
            self.font = FontConfig()
        if self.heading_font is None:
            self.heading_font = FontConfig(name="微软雅黑", size=14, bold=True)
        if self.code_font is None:
            self.code_font = FontConfig(name="Consolas", size=10)
        if self.page is None:
            self.page = PageConfig()


@dataclass
class ImageConfig:
    """图片配置"""
    max_width: int = 800
    max_height: int = 600
    quality: int = 85
    auto_resize: bool = True
    download_timeout: int = 30
    cache_enabled: bool = True
    cache_dir: str = "./cache/images"
    supported_formats: list = None
    
    def __post_init__(self):
        if self.supported_formats is None:
            self.supported_formats = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.tiff']


@dataclass
class ConversionConfig:
    """转换配置"""
    output_format: str = OutputFormat.DOCX.value
    output_dir: str = "./output"
    temp_dir: str = "./temp"
    backup_enabled: bool = True
    backup_dir: str = "./backup"
    overwrite_existing: bool = False
    preserve_structure: bool = True
    batch_size: int = 10
    max_workers: int = 4
    enable_parallel: bool = True
    retry_count: int = 3
    retry_delay: float = 1.0


@dataclass
class LogConfig:
    """日志配置"""
    level: str = LogLevel.INFO.value
    file_enabled: bool = True
    file_path: str = "./logs/converter.log"
    console_enabled: bool = True
    max_file_size: int = 10 * 1024 * 1024  # 10MB
    backup_count: int = 5
    format: str = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"


@dataclass
class AppConfig:
    """应用配置"""
    version: str = "1.0.0"
    debug: bool = False
    style: StyleConfig = None
    image: ImageConfig = None
    conversion: ConversionConfig = None
    log: LogConfig = None
    
    def __post_init__(self):
        if self.style is None:
            self.style = StyleConfig()
        if self.image is None:
            self.image = ImageConfig()
        if self.conversion is None:
            self.conversion = ConversionConfig()
        if self.log is None:
            self.log = LogConfig()


class ConfigManager:
    """配置管理器"""
    
    def __init__(self, config_dir: str = "./config"):
        """初始化配置管理器
        
        Args:
            config_dir: 配置文件目录
        """
        self.config_dir = Path(config_dir)
        self.config_file = self.config_dir / "config.yaml"
        self.user_config_file = self.config_dir / "user_config.yaml"
        self.themes_dir = self.config_dir / "themes"
        
        # 创建配置目录
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.themes_dir.mkdir(parents=True, exist_ok=True)
        
        # 默认配置
        self._default_config = AppConfig()
        self._current_config = None
        
        # 加载配置
        self.load_config()
    
    def load_config(self) -> AppConfig:
        """加载配置
        
        Returns:
            AppConfig: 应用配置对象
        """
        try:
            # 加载默认配置
            config_data = asdict(self._default_config)
            
            # 加载系统配置文件
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    system_config = yaml.safe_load(f) or {}
                config_data.update(system_config)
            
            # 加载用户配置文件
            if self.user_config_file.exists():
                with open(self.user_config_file, 'r', encoding='utf-8') as f:
                    user_config = yaml.safe_load(f) or {}
                self._merge_config(config_data, user_config)
            
            # 加载环境变量配置
            env_config = self._load_env_config()
            self._merge_config(config_data, env_config)
            
            # 创建配置对象
            self._current_config = self._dict_to_config(config_data)
            
            return self._current_config
            
        except Exception as e:
            print(f"加载配置失败: {e}")
            self._current_config = self._default_config
            return self._current_config
    
    def save_config(self, config: Optional[AppConfig] = None, user_config: bool = True) -> bool:
        """保存配置
        
        Args:
            config: 要保存的配置，如果为None则保存当前配置
            user_config: 是否保存为用户配置
            
        Returns:
            bool: 保存是否成功
        """
        try:
            if config is None:
                config = self._current_config
            
            config_data = asdict(config)
            
            # 选择保存文件
            target_file = self.user_config_file if user_config else self.config_file
            
            with open(target_file, 'w', encoding='utf-8') as f:
                yaml.dump(config_data, f, default_flow_style=False, 
                         allow_unicode=True, indent=2)
            
            return True
            
        except Exception as e:
            print(f"保存配置失败: {e}")
            return False
    
    def get_config(self) -> AppConfig:
        """获取当前配置
        
        Returns:
            AppConfig: 当前配置对象
        """
        if self._current_config is None:
            self.load_config()
        return self._current_config
    
    def update_config(self, **kwargs) -> bool:
        """更新配置
        
        Args:
            **kwargs: 要更新的配置项
            
        Returns:
            bool: 更新是否成功
        """
        try:
            if self._current_config is None:
                self.load_config()
            
            config_dict = asdict(self._current_config)
            self._merge_config(config_dict, kwargs)
            
            self._current_config = self._dict_to_config(config_dict)
            
            return True
            
        except Exception as e:
            print(f"更新配置失败: {e}")
            return False
    
    def reset_config(self) -> bool:
        """重置配置为默认值
        
        Returns:
            bool: 重置是否成功
        """
        try:
            self._current_config = AppConfig()
            return self.save_config()
        except Exception as e:
            print(f"重置配置失败: {e}")
            return False
    
    def load_theme(self, theme_name: str) -> Optional[Dict[str, Any]]:
        """加载主题配置
        
        Args:
            theme_name: 主题名称
            
        Returns:
            Optional[Dict[str, Any]]: 主题配置字典
        """
        try:
            theme_file = self.themes_dir / f"{theme_name}.yaml"
            
            if not theme_file.exists():
                # 创建默认主题文件
                self._create_default_themes()
            
            if theme_file.exists():
                with open(theme_file, 'r', encoding='utf-8') as f:
                    return yaml.safe_load(f)
            
            return None
            
        except Exception as e:
            print(f"加载主题失败 {theme_name}: {e}")
            return None
    
    def save_theme(self, theme_name: str, theme_config: Dict[str, Any]) -> bool:
        """保存主题配置
        
        Args:
            theme_name: 主题名称
            theme_config: 主题配置
            
        Returns:
            bool: 保存是否成功
        """
        try:
            theme_file = self.themes_dir / f"{theme_name}.yaml"
            
            with open(theme_file, 'w', encoding='utf-8') as f:
                yaml.dump(theme_config, f, default_flow_style=False, 
                         allow_unicode=True, indent=2)
            
            return True
            
        except Exception as e:
            print(f"保存主题失败 {theme_name}: {e}")
            return False
    
    def list_themes(self) -> list:
        """列出所有可用主题
        
        Returns:
            list: 主题名称列表
        """
        try:
            themes = []
            
            if self.themes_dir.exists():
                for theme_file in self.themes_dir.glob("*.yaml"):
                    themes.append(theme_file.stem)
            
            # 如果没有主题文件，创建默认主题
            if not themes:
                self._create_default_themes()
                themes = [theme.value for theme in StyleTheme]
            
            return sorted(themes)
            
        except Exception as e:
            print(f"列出主题失败: {e}")
            return []
    
    def apply_theme(self, theme_name: str) -> bool:
        """应用主题
        
        Args:
            theme_name: 主题名称
            
        Returns:
            bool: 应用是否成功
        """
        try:
            theme_config = self.load_theme(theme_name)
            
            if theme_config:
                # 更新样式配置
                if 'style' in theme_config:
                    style_dict = asdict(self._current_config.style)
                    self._merge_config(style_dict, theme_config['style'])
                    self._current_config.style = self._dict_to_style_config(style_dict)
                
                return True
            
            return False
            
        except Exception as e:
            print(f"应用主题失败 {theme_name}: {e}")
            return False
    
    def export_config(self, file_path: str, format: str = "yaml") -> bool:
        """导出配置
        
        Args:
            file_path: 导出文件路径
            format: 导出格式（yaml/json）
            
        Returns:
            bool: 导出是否成功
        """
        try:
            config_data = asdict(self.get_config())
            
            with open(file_path, 'w', encoding='utf-8') as f:
                if format.lower() == 'json':
                    json.dump(config_data, f, indent=2, ensure_ascii=False)
                else:
                    yaml.dump(config_data, f, default_flow_style=False, 
                             allow_unicode=True, indent=2)
            
            return True
            
        except Exception as e:
            print(f"导出配置失败: {e}")
            return False
    
    def import_config(self, file_path: str) -> bool:
        """导入配置
        
        Args:
            file_path: 配置文件路径
            
        Returns:
            bool: 导入是否成功
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                if file_path.endswith('.json'):
                    config_data = json.load(f)
                else:
                    config_data = yaml.safe_load(f)
            
            self._current_config = self._dict_to_config(config_data)
            
            return self.save_config()
            
        except Exception as e:
            print(f"导入配置失败: {e}")
            return False
    
    def _merge_config(self, base: Dict[str, Any], update: Dict[str, Any]):
        """递归合并配置字典"""
        for key, value in update.items():
            if key in base and isinstance(base[key], dict) and isinstance(value, dict):
                self._merge_config(base[key], value)
            else:
                base[key] = value
    
    def _load_env_config(self) -> Dict[str, Any]:
        """从环境变量加载配置"""
        env_config = {}
        
        # 定义环境变量映射
        env_mappings = {
            'MD2WORD_DEBUG': ('debug', bool),
            'MD2WORD_OUTPUT_DIR': ('conversion.output_dir', str),
            'MD2WORD_TEMP_DIR': ('conversion.temp_dir', str),
            'MD2WORD_LOG_LEVEL': ('log.level', str),
            'MD2WORD_MAX_WORKERS': ('conversion.max_workers', int),
            'MD2WORD_THEME': ('style.theme', str),
        }
        
        for env_var, (config_path, config_type) in env_mappings.items():
            env_value = os.getenv(env_var)
            if env_value is not None:
                try:
                    # 类型转换
                    if config_type == bool:
                        value = env_value.lower() in ('true', '1', 'yes', 'on')
                    elif config_type == int:
                        value = int(env_value)
                    else:
                        value = env_value
                    
                    # 设置嵌套配置
                    self._set_nested_config(env_config, config_path, value)
                    
                except ValueError:
                    print(f"环境变量 {env_var} 值无效: {env_value}")
        
        return env_config
    
    def _set_nested_config(self, config: Dict[str, Any], path: str, value: Any):
        """设置嵌套配置值"""
        keys = path.split('.')
        current = config
        
        for key in keys[:-1]:
            if key not in current:
                current[key] = {}
            current = current[key]
        
        current[keys[-1]] = value
    
    def _dict_to_config(self, config_dict: Dict[str, Any]) -> AppConfig:
        """将字典转换为配置对象"""
        # 处理嵌套对象
        if 'style' in config_dict:
            config_dict['style'] = self._dict_to_style_config(config_dict['style'])
        if 'image' in config_dict:
            config_dict['image'] = ImageConfig(**config_dict['image'])
        if 'conversion' in config_dict:
            config_dict['conversion'] = ConversionConfig(**config_dict['conversion'])
        if 'log' in config_dict:
            config_dict['log'] = LogConfig(**config_dict['log'])
        
        return AppConfig(**config_dict)
    
    def _dict_to_style_config(self, style_dict: Dict[str, Any]) -> StyleConfig:
        """将字典转换为样式配置对象"""
        if 'font' in style_dict:
            style_dict['font'] = FontConfig(**style_dict['font'])
        if 'heading_font' in style_dict:
            style_dict['heading_font'] = FontConfig(**style_dict['heading_font'])
        if 'code_font' in style_dict:
            style_dict['code_font'] = FontConfig(**style_dict['code_font'])
        if 'page' in style_dict:
            style_dict['page'] = PageConfig(**style_dict['page'])
        
        return StyleConfig(**style_dict)
    
    def _create_default_themes(self):
        """创建默认主题文件"""
        themes = {
            'default': {
                'style': {
                    'theme': 'default',
                    'font': {'name': '微软雅黑', 'size': 12},
                    'heading_font': {'name': '微软雅黑', 'size': 14, 'bold': True},
                    'line_spacing': 1.15
                }
            },
            'academic': {
                'style': {
                    'theme': 'academic',
                    'font': {'name': 'Times New Roman', 'size': 12},
                    'heading_font': {'name': 'Times New Roman', 'size': 14, 'bold': True},
                    'line_spacing': 1.5,
                    'enable_toc': True
                }
            },
            'business': {
                'style': {
                    'theme': 'business',
                    'font': {'name': 'Calibri', 'size': 11},
                    'heading_font': {'name': 'Calibri', 'size': 14, 'bold': True},
                    'line_spacing': 1.15,
                    'enable_header': True,
                    'enable_footer': True
                }
            }
        }
        
        for theme_name, theme_config in themes.items():
            self.save_theme(theme_name, theme_config)


# 全局配置管理器实例
_config_manager = None


def get_config_manager() -> ConfigManager:
    """获取全局配置管理器实例
    
    Returns:
        ConfigManager: 配置管理器实例
    """
    global _config_manager
    if _config_manager is None:
        _config_manager = ConfigManager()
    return _config_manager


def get_config() -> AppConfig:
    """获取当前应用配置
    
    Returns:
        AppConfig: 应用配置对象
    """
    return get_config_manager().get_config()


def update_config(**kwargs) -> bool:
    """更新应用配置
    
    Args:
        **kwargs: 要更新的配置项
        
    Returns:
        bool: 更新是否成功
    """
    return get_config_manager().update_config(**kwargs)


def save_config(config: Optional[AppConfig] = None) -> bool:
    """保存应用配置
    
    Args:
        config: 要保存的配置
        
    Returns:
        bool: 保存是否成功
    """
    return get_config_manager().save_config(config)