#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件处理工具模块
负责文件的读取、写入、路径处理等操作
"""

import os
import shutil
import mimetypes
from pathlib import Path
from typing import Optional, List, Dict, Any, Union
from urllib.parse import urlparse
from urllib.request import urlretrieve

try:
    from docx import Document
except ImportError:
    raise ImportError("请安装python-docx库: pip install python-docx")


class FileHandler:
    """文件处理器"""
    
    def __init__(self, encoding: str = 'utf-8'):
        """初始化文件处理器
        
        Args:
            encoding: 默认文件编码
        """
        self.encoding = encoding
        self.supported_image_formats = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.webp'}
        self.supported_markdown_formats = {'.md', '.markdown', '.mdown', '.mkd'}
    
    def read_file(self, file_path: str, encoding: Optional[str] = None) -> str:
        """读取文本文件
        
        Args:
            file_path: 文件路径
            encoding: 文件编码，默认使用初始化时的编码
            
        Returns:
            str: 文件内容
            
        Raises:
            FileNotFoundError: 文件不存在
            UnicodeDecodeError: 编码错误
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        file_encoding = encoding or self.encoding
        
        try:
            with open(file_path, 'r', encoding=file_encoding) as f:
                return f.read()
        except UnicodeDecodeError:
            # 尝试其他常见编码
            for enc in ['utf-8', 'gbk', 'gb2312', 'latin-1']:
                if enc != file_encoding:
                    try:
                        with open(file_path, 'r', encoding=enc) as f:
                            content = f.read()
                            print(f"警告: 使用 {enc} 编码读取文件 {file_path}")
                            return content
                    except UnicodeDecodeError:
                        continue
            raise UnicodeDecodeError(f"无法解码文件 {file_path}")
    
    def write_file(self, file_path: str, content: str, encoding: Optional[str] = None) -> bool:
        """写入文本文件
        
        Args:
            file_path: 文件路径
            content: 文件内容
            encoding: 文件编码，默认使用初始化时的编码
            
        Returns:
            bool: 写入是否成功
        """
        try:
            # 确保目录存在
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            
            file_encoding = encoding or self.encoding
            with open(file_path, 'w', encoding=file_encoding) as f:
                f.write(content)
            return True
        except Exception as e:
            print(f"写入文件失败 {file_path}: {e}")
            return False
    
    def copy_file(self, src_path: str, dst_path: str) -> bool:
        """复制文件
        
        Args:
            src_path: 源文件路径
            dst_path: 目标文件路径
            
        Returns:
            bool: 复制是否成功
        """
        try:
            if not os.path.exists(src_path):
                print(f"源文件不存在: {src_path}")
                return False
            
            # 确保目标目录存在
            os.makedirs(os.path.dirname(dst_path), exist_ok=True)
            
            shutil.copy2(src_path, dst_path)
            return True
        except Exception as e:
            print(f"复制文件失败 {src_path} -> {dst_path}: {e}")
            return False
    
    def move_file(self, src_path: str, dst_path: str) -> bool:
        """移动文件
        
        Args:
            src_path: 源文件路径
            dst_path: 目标文件路径
            
        Returns:
            bool: 移动是否成功
        """
        try:
            if not os.path.exists(src_path):
                print(f"源文件不存在: {src_path}")
                return False
            
            # 确保目标目录存在
            os.makedirs(os.path.dirname(dst_path), exist_ok=True)
            
            shutil.move(src_path, dst_path)
            return True
        except Exception as e:
            print(f"移动文件失败 {src_path} -> {dst_path}: {e}")
            return False
    
    def delete_file(self, file_path: str) -> bool:
        """删除文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            bool: 删除是否成功
        """
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
            return True
        except Exception as e:
            print(f"删除文件失败 {file_path}: {e}")
            return False
    
    def create_directory(self, dir_path: str) -> bool:
        """创建目录
        
        Args:
            dir_path: 目录路径
            
        Returns:
            bool: 创建是否成功
        """
        try:
            os.makedirs(dir_path, exist_ok=True)
            return True
        except Exception as e:
            print(f"创建目录失败 {dir_path}: {e}")
            return False
    
    def list_files(self, dir_path: str, pattern: Optional[str] = None, 
                   recursive: bool = False) -> List[str]:
        """列出目录中的文件
        
        Args:
            dir_path: 目录路径
            pattern: 文件名模式（支持通配符）
            recursive: 是否递归搜索子目录
            
        Returns:
            List[str]: 文件路径列表
        """
        if not os.path.exists(dir_path):
            return []
        
        files = []
        
        if recursive:
            for root, dirs, filenames in os.walk(dir_path):
                for filename in filenames:
                    file_path = os.path.join(root, filename)
                    if pattern is None or self._match_pattern(filename, pattern):
                        files.append(file_path)
        else:
            for item in os.listdir(dir_path):
                item_path = os.path.join(dir_path, item)
                if os.path.isfile(item_path):
                    if pattern is None or self._match_pattern(item, pattern):
                        files.append(item_path)
        
        return sorted(files)
    
    def find_markdown_files(self, dir_path: str, recursive: bool = True) -> List[str]:
        """查找Markdown文件
        
        Args:
            dir_path: 目录路径
            recursive: 是否递归搜索
            
        Returns:
            List[str]: Markdown文件路径列表
        """
        markdown_files = []
        
        if not os.path.exists(dir_path):
            return markdown_files
        
        if recursive:
            for root, dirs, files in os.walk(dir_path):
                for file in files:
                    if self.is_markdown_file(file):
                        markdown_files.append(os.path.join(root, file))
        else:
            for item in os.listdir(dir_path):
                item_path = os.path.join(dir_path, item)
                if os.path.isfile(item_path) and self.is_markdown_file(item):
                    markdown_files.append(item_path)
        
        return sorted(markdown_files)
    
    def is_markdown_file(self, file_path: str) -> bool:
        """判断是否为Markdown文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            bool: 是否为Markdown文件
        """
        _, ext = os.path.splitext(file_path.lower())
        return ext in self.supported_markdown_formats
    
    def is_image_file(self, file_path: str) -> bool:
        """判断是否为图片文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            bool: 是否为图片文件
        """
        _, ext = os.path.splitext(file_path.lower())
        return ext in self.supported_image_formats
    
    def get_file_info(self, file_path: str) -> Dict[str, Any]:
        """获取文件信息
        
        Args:
            file_path: 文件路径
            
        Returns:
            Dict[str, Any]: 文件信息字典
        """
        if not os.path.exists(file_path):
            return {}
        
        stat = os.stat(file_path)
        
        return {
            'path': file_path,
            'name': os.path.basename(file_path),
            'size': stat.st_size,
            'created': stat.st_ctime,
            'modified': stat.st_mtime,
            'is_file': os.path.isfile(file_path),
            'is_dir': os.path.isdir(file_path),
            'extension': os.path.splitext(file_path)[1].lower(),
            'mime_type': mimetypes.guess_type(file_path)[0]
        }
    
    def get_relative_path(self, file_path: str, base_path: str) -> str:
        """获取相对路径
        
        Args:
            file_path: 文件路径
            base_path: 基准路径
            
        Returns:
            str: 相对路径
        """
        return os.path.relpath(file_path, base_path)
    
    def resolve_path(self, path: str, base_path: Optional[str] = None) -> str:
        """解析路径（处理相对路径、绝对路径等）
        
        Args:
            path: 路径
            base_path: 基准路径
            
        Returns:
            str: 解析后的绝对路径
        """
        if os.path.isabs(path):
            return os.path.normpath(path)
        
        if base_path:
            return os.path.normpath(os.path.join(base_path, path))
        
        return os.path.normpath(os.path.abspath(path))
    
    def download_file(self, url: str, local_path: str) -> bool:
        """下载文件
        
        Args:
            url: 文件URL
            local_path: 本地保存路径
            
        Returns:
            bool: 下载是否成功
        """
        try:
            # 确保目录存在
            os.makedirs(os.path.dirname(local_path), exist_ok=True)
            
            urlretrieve(url, local_path)
            return True
        except Exception as e:
            print(f"下载文件失败 {url}: {e}")
            return False
    
    def is_url(self, path: str) -> bool:
        """判断是否为URL
        
        Args:
            path: 路径字符串
            
        Returns:
            bool: 是否为URL
        """
        try:
            result = urlparse(path)
            return all([result.scheme, result.netloc])
        except:
            return False
    
    def save_document(self, document: Document, file_path: str) -> bool:
        """保存Word文档
        
        Args:
            document: Word文档对象
            file_path: 保存路径
            
        Returns:
            bool: 保存是否成功
        """
        try:
            # 确保目录存在
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            
            document.save(file_path)
            return True
        except Exception as e:
            print(f"保存Word文档失败 {file_path}: {e}")
            return False
    
    def load_document(self, file_path: str) -> Optional[Document]:
        """加载Word文档
        
        Args:
            file_path: 文档路径
            
        Returns:
            Optional[Document]: Word文档对象，失败时返回None
        """
        try:
            if not os.path.exists(file_path):
                print(f"文档文件不存在: {file_path}")
                return None
            
            return Document(file_path)
        except Exception as e:
            print(f"加载Word文档失败 {file_path}: {e}")
            return None
    
    def backup_file(self, file_path: str, backup_suffix: str = '.bak') -> bool:
        """备份文件
        
        Args:
            file_path: 文件路径
            backup_suffix: 备份文件后缀
            
        Returns:
            bool: 备份是否成功
        """
        if not os.path.exists(file_path):
            return False
        
        backup_path = file_path + backup_suffix
        return self.copy_file(file_path, backup_path)
    
    def clean_filename(self, filename: str) -> str:
        """清理文件名（移除非法字符）
        
        Args:
            filename: 原始文件名
            
        Returns:
            str: 清理后的文件名
        """
        # 移除或替换非法字符
        illegal_chars = '<>:"/\\|?*'
        for char in illegal_chars:
            filename = filename.replace(char, '_')
        
        # 移除前后空格和点
        filename = filename.strip(' .')
        
        # 确保文件名不为空
        if not filename:
            filename = 'untitled'
        
        return filename
    
    def get_unique_filename(self, file_path: str) -> str:
        """获取唯一文件名（如果文件已存在，添加数字后缀）
        
        Args:
            file_path: 文件路径
            
        Returns:
            str: 唯一文件路径
        """
        if not os.path.exists(file_path):
            return file_path
        
        base, ext = os.path.splitext(file_path)
        counter = 1
        
        while True:
            new_path = f"{base}_{counter}{ext}"
            if not os.path.exists(new_path):
                return new_path
            counter += 1
    
    def _match_pattern(self, filename: str, pattern: str) -> bool:
        """匹配文件名模式
        
        Args:
            filename: 文件名
            pattern: 模式字符串
            
        Returns:
            bool: 是否匹配
        """
        import fnmatch
        return fnmatch.fnmatch(filename.lower(), pattern.lower())
    
    def get_file_size_human(self, file_path: str) -> str:
        """获取人类可读的文件大小
        
        Args:
            file_path: 文件路径
            
        Returns:
            str: 人类可读的文件大小
        """
        if not os.path.exists(file_path):
            return "0 B"
        
        size = os.path.getsize(file_path)
        
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        
        return f"{size:.1f} PB"
    
    def validate_path(self, path: str) -> Dict[str, Any]:
        """验证路径
        
        Args:
            path: 路径字符串
            
        Returns:
            Dict[str, Any]: 验证结果
        """
        result = {
            'valid': True,
            'exists': False,
            'is_file': False,
            'is_dir': False,
            'readable': False,
            'writable': False,
            'errors': []
        }
        
        try:
            # 检查路径格式
            if not path or not isinstance(path, str):
                result['valid'] = False
                result['errors'].append('路径为空或格式无效')
                return result
            
            # 检查是否存在
            if os.path.exists(path):
                result['exists'] = True
                result['is_file'] = os.path.isfile(path)
                result['is_dir'] = os.path.isdir(path)
                result['readable'] = os.access(path, os.R_OK)
                result['writable'] = os.access(path, os.W_OK)
            else:
                # 检查父目录是否存在且可写
                parent_dir = os.path.dirname(path)
                if parent_dir and os.path.exists(parent_dir):
                    result['writable'] = os.access(parent_dir, os.W_OK)
                else:
                    result['errors'].append('父目录不存在')
            
        except Exception as e:
            result['valid'] = False
            result['errors'].append(f'路径验证失败: {e}')
        
        return result