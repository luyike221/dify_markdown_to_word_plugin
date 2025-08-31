#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown to Word 转换器主程序
提供命令行接口和API接口
"""

import argparse
import sys
import os
from pathlib import Path
from typing import Optional, Dict, Any

# 添加src目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from converters.markdown_parser import MarkdownParser
from converters.word_generator import WordGenerator
from converters.style_handler import StyleHandler
from utils.config import Config
from utils.file_handler import FileHandler


class MarkdownToWordConverter:
    """Markdown到Word转换器主类"""
    
    def __init__(self, config_path: Optional[str] = None, style_config_path: Optional[str] = None):
        """初始化转换器
        
        Args:
            config_path: 配置文件路径
            style_config_path: 样式配置文件路径
        """
        self.config = Config(config_path)
        self.style_handler = StyleHandler(style_config_path)
        self.markdown_parser = MarkdownParser()
        self.word_generator = WordGenerator(self.style_handler)
        self.file_handler = FileHandler()
    
    def convert_file(self, input_path: str, output_path: str, 
                    theme: str = 'default', **kwargs) -> bool:
        """转换单个文件
        
        Args:
            input_path: 输入Markdown文件路径
            output_path: 输出Word文件路径
            theme: 样式主题名称
            **kwargs: 其他转换选项
            
        Returns:
            bool: 转换是否成功
        """
        try:
            # 验证输入文件
            if not os.path.exists(input_path):
                print(f"错误: 输入文件不存在: {input_path}")
                return False
            
            if not input_path.lower().endswith(('.md', '.markdown')):
                print(f"警告: 输入文件可能不是Markdown格式: {input_path}")
            
            # 读取Markdown文件
            print(f"正在读取文件: {input_path}")
            markdown_content = self.file_handler.read_file(input_path)
            
            # 解析Markdown
            print("正在解析Markdown内容...")
            parsed_content = self.markdown_parser.parse(markdown_content)
            
            # 设置样式主题
            if not self.style_handler.set_theme(theme):
                print(f"警告: 主题 '{theme}' 不存在，使用默认主题")
                self.style_handler.set_theme('default')
            
            # 生成Word文档
            print("正在生成Word文档...")
            doc = self.word_generator.generate(parsed_content, **kwargs)
            
            # 保存文档
            print(f"正在保存文档: {output_path}")
            self.file_handler.save_document(doc, output_path)
            
            print(f"转换完成: {input_path} -> {output_path}")
            return True
            
        except Exception as e:
            print(f"转换失败: {e}")
            return False
    
    def convert_directory(self, input_dir: str, output_dir: str, 
                         theme: str = 'default', **kwargs) -> Dict[str, bool]:
        """批量转换目录中的Markdown文件
        
        Args:
            input_dir: 输入目录路径
            output_dir: 输出目录路径
            theme: 样式主题名称
            **kwargs: 其他转换选项
            
        Returns:
            Dict[str, bool]: 文件转换结果字典
        """
        results = {}
        
        if not os.path.exists(input_dir):
            print(f"错误: 输入目录不存在: {input_dir}")
            return results
        
        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)
        
        # 查找所有Markdown文件
        markdown_files = []
        for root, dirs, files in os.walk(input_dir):
            for file in files:
                if file.lower().endswith(('.md', '.markdown')):
                    markdown_files.append(os.path.join(root, file))
        
        if not markdown_files:
            print(f"在目录 {input_dir} 中未找到Markdown文件")
            return results
        
        print(f"找到 {len(markdown_files)} 个Markdown文件")
        
        # 批量转换
        for md_file in markdown_files:
            # 计算相对路径
            rel_path = os.path.relpath(md_file, input_dir)
            # 生成输出文件路径
            output_file = os.path.join(output_dir, 
                                     os.path.splitext(rel_path)[0] + '.docx')
            
            # 创建输出文件的目录
            os.makedirs(os.path.dirname(output_file), exist_ok=True)
            
            # 转换文件
            success = self.convert_file(md_file, output_file, theme, **kwargs)
            results[md_file] = success
        
        # 输出统计信息
        success_count = sum(1 for success in results.values() if success)
        print(f"\n批量转换完成: {success_count}/{len(markdown_files)} 个文件转换成功")
        
        return results
    
    def convert_text(self, markdown_text: str, output_path: str, 
                    theme: str = 'default', **kwargs) -> bool:
        """转换Markdown文本
        
        Args:
            markdown_text: Markdown文本内容
            output_path: 输出Word文件路径
            theme: 样式主题名称
            **kwargs: 其他转换选项
            
        Returns:
            bool: 转换是否成功
        """
        try:
            # 解析Markdown
            parsed_content = self.markdown_parser.parse(markdown_text)
            
            # 设置样式主题
            if not self.style_handler.set_theme(theme):
                print(f"警告: 主题 '{theme}' 不存在，使用默认主题")
                self.style_handler.set_theme('default')
            
            # 生成Word文档
            doc = self.word_generator.generate(parsed_content, **kwargs)
            
            # 保存文档
            self.file_handler.save_document(doc, output_path)
            
            return True
            
        except Exception as e:
            print(f"转换失败: {e}")
            return False
    
    def list_themes(self) -> list:
        """列出可用的样式主题"""
        return self.style_handler.list_themes()
    
    def get_theme_preview(self, theme_name: str) -> Dict[str, Any]:
        """获取主题预览信息"""
        if self.style_handler.set_theme(theme_name):
            preview = {}
            for element_type in ['normal', 'heading1', 'heading2', 'code_block', 'quote']:
                preview[element_type] = self.style_handler.get_style_preview(element_type)
            return preview
        return {}


def create_parser() -> argparse.ArgumentParser:
    """创建命令行参数解析器"""
    parser = argparse.ArgumentParser(
        description='Markdown to Word 转换器',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法:
  %(prog)s input.md output.docx
  %(prog)s input.md output.docx --theme academic
  %(prog)s --input-dir ./docs --output-dir ./output
  %(prog)s --list-themes
        """
    )
    
    # 输入输出参数
    parser.add_argument('input', nargs='?', help='输入Markdown文件路径')
    parser.add_argument('output', nargs='?', help='输出Word文件路径')
    
    # 目录批量转换
    parser.add_argument('--input-dir', help='输入目录路径（批量转换）')
    parser.add_argument('--output-dir', help='输出目录路径（批量转换）')
    
    # 样式选项
    parser.add_argument('--theme', default='default', 
                       help='样式主题 (default: %(default)s)')
    parser.add_argument('--style-config', help='样式配置文件路径')
    
    # 转换选项
    parser.add_argument('--template', help='Word模板文件路径')
    parser.add_argument('--title', help='文档标题')
    parser.add_argument('--author', help='文档作者')
    parser.add_argument('--subject', help='文档主题')
    parser.add_argument('--keywords', help='文档关键词')
    
    # 功能选项
    parser.add_argument('--list-themes', action='store_true', 
                       help='列出可用的样式主题')
    parser.add_argument('--preview-theme', help='预览指定主题的样式')
    
    # 其他选项
    parser.add_argument('--config', help='配置文件路径')
    parser.add_argument('--verbose', '-v', action='store_true', 
                       help='显示详细信息')
    parser.add_argument('--version', action='version', version='%(prog)s 1.0.0')
    
    return parser


def main():
    """主函数"""
    parser = create_parser()
    args = parser.parse_args()
    
    # 创建转换器实例
    converter = MarkdownToWordConverter(
        config_path=args.config,
        style_config_path=args.style_config
    )
    
    # 列出主题
    if args.list_themes:
        themes = converter.list_themes()
        print("可用的样式主题:")
        for theme in themes:
            print(f"  - {theme}")
        return
    
    # 预览主题
    if args.preview_theme:
        preview = converter.get_theme_preview(args.preview_theme)
        if preview:
            print(f"主题 '{args.preview_theme}' 预览:")
            for element_type, style_info in preview.items():
                print(f"\n{element_type}:")
                if style_info.get('font'):
                    font = style_info['font']
                    print(f"  字体: {font.get('name', 'N/A')} {font.get('size', 'N/A')}pt")
                    if font.get('bold'):
                        print("  粗体: 是")
                    if font.get('italic'):
                        print("  斜体: 是")
        else:
            print(f"主题 '{args.preview_theme}' 不存在")
        return
    
    # 批量转换
    if args.input_dir and args.output_dir:
        kwargs = {}
        if args.template:
            kwargs['template_path'] = args.template
        if args.title:
            kwargs['title'] = args.title
        if args.author:
            kwargs['author'] = args.author
        if args.subject:
            kwargs['subject'] = args.subject
        if args.keywords:
            kwargs['keywords'] = args.keywords
        
        results = converter.convert_directory(
            args.input_dir, args.output_dir, args.theme, **kwargs
        )
        
        # 显示失败的文件
        failed_files = [f for f, success in results.items() if not success]
        if failed_files:
            print("\n转换失败的文件:")
            for file in failed_files:
                print(f"  - {file}")
        
        return
    
    # 单文件转换
    if args.input and args.output:
        kwargs = {}
        if args.template:
            kwargs['template_path'] = args.template
        if args.title:
            kwargs['title'] = args.title
        if args.author:
            kwargs['author'] = args.author
        if args.subject:
            kwargs['subject'] = args.subject
        if args.keywords:
            kwargs['keywords'] = args.keywords
        
        success = converter.convert_file(
            args.input, args.output, args.theme, **kwargs
        )
        
        if not success:
            sys.exit(1)
        return
    
    # 如果没有提供有效参数，显示帮助
    parser.print_help()


if __name__ == '__main__':
    main()