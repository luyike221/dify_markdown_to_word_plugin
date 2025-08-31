#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图片处理工具模块
负责图片的下载、处理、格式转换等操作
"""

import os
import io
import base64
import hashlib
from typing import Optional, Tuple, Dict, Any, Union
from urllib.parse import urlparse
from urllib.request import urlopen, Request

try:
    from PIL import Image, ImageOps
except ImportError:
    raise ImportError("请安装Pillow库: pip install Pillow")

try:
    from docx.shared import Inches, Cm
except ImportError:
    raise ImportError("请安装python-docx库: pip install python-docx")


class ImageProcessor:
    """图片处理器"""
    
    def __init__(self, cache_dir: str = "./cache/images", max_width: int = 800, max_height: int = 600):
        """初始化图片处理器
        
        Args:
            cache_dir: 图片缓存目录
            max_width: 最大宽度（像素）
            max_height: 最大高度（像素）
        """
        self.cache_dir = cache_dir
        self.max_width = max_width
        self.max_height = max_height
        self.supported_formats = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.tiff', '.svg'}
        
        # 创建缓存目录
        os.makedirs(cache_dir, exist_ok=True)
    
    def process_image(self, image_path: str, output_path: Optional[str] = None, 
                     resize: bool = True, quality: int = 85) -> Optional[str]:
        """处理图片
        
        Args:
            image_path: 图片路径（本地路径或URL）
            output_path: 输出路径，如果为None则保存到缓存目录
            resize: 是否调整大小
            quality: 图片质量（1-100）
            
        Returns:
            Optional[str]: 处理后的图片路径，失败时返回None
        """
        try:
            # 获取图片
            if self._is_url(image_path):
                image_data = self._download_image(image_path)
                if not image_data:
                    return None
                image = Image.open(io.BytesIO(image_data))
                
                # 生成缓存文件名
                if not output_path:
                    filename = self._generate_cache_filename(image_path)
                    output_path = os.path.join(self.cache_dir, filename)
            else:
                if not os.path.exists(image_path):
                    print(f"图片文件不存在: {image_path}")
                    return None
                
                image = Image.open(image_path)
                
                if not output_path:
                    output_path = image_path
            
            # 处理图片
            processed_image = self._process_image_object(image, resize)
            
            # 保存图片
            self._save_image(processed_image, output_path, quality)
            
            return output_path
            
        except Exception as e:
            print(f"处理图片失败 {image_path}: {e}")
            return None
    
    def get_image_info(self, image_path: str) -> Dict[str, Any]:
        """获取图片信息
        
        Args:
            image_path: 图片路径
            
        Returns:
            Dict[str, Any]: 图片信息字典
        """
        try:
            if self._is_url(image_path):
                image_data = self._download_image(image_path)
                if not image_data:
                    return {}
                image = Image.open(io.BytesIO(image_data))
            else:
                if not os.path.exists(image_path):
                    return {}
                image = Image.open(image_path)
            
            return {
                'path': image_path,
                'format': image.format,
                'mode': image.mode,
                'size': image.size,
                'width': image.width,
                'height': image.height,
                'has_transparency': self._has_transparency(image),
                'file_size': len(image_data) if self._is_url(image_path) else os.path.getsize(image_path)
            }
            
        except Exception as e:
            print(f"获取图片信息失败 {image_path}: {e}")
            return {}
    
    def resize_image(self, image_path: str, width: Optional[int] = None, 
                    height: Optional[int] = None, maintain_aspect: bool = True) -> Optional[str]:
        """调整图片大小
        
        Args:
            image_path: 图片路径
            width: 目标宽度
            height: 目标高度
            maintain_aspect: 是否保持宽高比
            
        Returns:
            Optional[str]: 调整后的图片路径
        """
        try:
            if self._is_url(image_path):
                image_data = self._download_image(image_path)
                if not image_data:
                    return None
                image = Image.open(io.BytesIO(image_data))
                output_path = os.path.join(self.cache_dir, self._generate_cache_filename(image_path))
            else:
                if not os.path.exists(image_path):
                    return None
                image = Image.open(image_path)
                output_path = image_path
            
            # 计算新尺寸
            new_size = self._calculate_new_size(image.size, width, height, maintain_aspect)
            
            # 调整大小
            resized_image = image.resize(new_size, Image.Resampling.LANCZOS)
            
            # 保存
            self._save_image(resized_image, output_path)
            
            return output_path
            
        except Exception as e:
            print(f"调整图片大小失败 {image_path}: {e}")
            return None
    
    def convert_format(self, image_path: str, target_format: str, 
                      output_path: Optional[str] = None) -> Optional[str]:
        """转换图片格式
        
        Args:
            image_path: 图片路径
            target_format: 目标格式（如'PNG', 'JPEG'等）
            output_path: 输出路径
            
        Returns:
            Optional[str]: 转换后的图片路径
        """
        try:
            if self._is_url(image_path):
                image_data = self._download_image(image_path)
                if not image_data:
                    return None
                image = Image.open(io.BytesIO(image_data))
            else:
                if not os.path.exists(image_path):
                    return None
                image = Image.open(image_path)
            
            # 生成输出路径
            if not output_path:
                base_name = os.path.splitext(os.path.basename(image_path))[0]
                ext = '.jpg' if target_format.upper() == 'JPEG' else f'.{target_format.lower()}'
                output_path = os.path.join(self.cache_dir, f"{base_name}{ext}")
            
            # 处理透明度
            if target_format.upper() == 'JPEG' and image.mode in ('RGBA', 'LA', 'P'):
                # JPEG不支持透明度，添加白色背景
                background = Image.new('RGB', image.size, (255, 255, 255))
                if image.mode == 'P':
                    image = image.convert('RGBA')
                background.paste(image, mask=image.split()[-1] if image.mode == 'RGBA' else None)
                image = background
            
            # 保存
            image.save(output_path, format=target_format.upper())
            
            return output_path
            
        except Exception as e:
            print(f"转换图片格式失败 {image_path}: {e}")
            return None
    
    def optimize_for_word(self, image_path: str, max_width_inches: float = 6.0, 
                         max_height_inches: float = 8.0, dpi: int = 96) -> Optional[str]:
        """为Word文档优化图片
        
        Args:
            image_path: 图片路径
            max_width_inches: 最大宽度（英寸）
            max_height_inches: 最大高度（英寸）
            dpi: 分辨率
            
        Returns:
            Optional[str]: 优化后的图片路径
        """
        try:
            # 计算像素尺寸
            max_width_px = int(max_width_inches * dpi)
            max_height_px = int(max_height_inches * dpi)
            
            # 处理图片
            return self.resize_image(image_path, max_width_px, max_height_px, True)
            
        except Exception as e:
            print(f"优化图片失败 {image_path}: {e}")
            return None
    
    def get_word_dimensions(self, image_path: str, max_width_inches: float = 6.0, 
                           max_height_inches: float = 8.0) -> Tuple[float, float]:
        """获取适合Word文档的图片尺寸
        
        Args:
            image_path: 图片路径
            max_width_inches: 最大宽度（英寸）
            max_height_inches: 最大高度（英寸）
            
        Returns:
            Tuple[float, float]: (宽度英寸, 高度英寸)
        """
        try:
            info = self.get_image_info(image_path)
            if not info:
                return (max_width_inches, max_height_inches)
            
            width, height = info['size']
            
            # 计算缩放比例
            width_ratio = max_width_inches / (width / 96)  # 假设96 DPI
            height_ratio = max_height_inches / (height / 96)
            
            scale = min(width_ratio, height_ratio, 1.0)  # 不放大
            
            final_width = (width / 96) * scale
            final_height = (height / 96) * scale
            
            return (final_width, final_height)
            
        except Exception as e:
            print(f"计算Word尺寸失败 {image_path}: {e}")
            return (max_width_inches, max_height_inches)
    
    def create_thumbnail(self, image_path: str, size: Tuple[int, int] = (150, 150)) -> Optional[str]:
        """创建缩略图
        
        Args:
            image_path: 图片路径
            size: 缩略图尺寸
            
        Returns:
            Optional[str]: 缩略图路径
        """
        try:
            if self._is_url(image_path):
                image_data = self._download_image(image_path)
                if not image_data:
                    return None
                image = Image.open(io.BytesIO(image_data))
                base_name = self._generate_cache_filename(image_path, suffix='_thumb')
            else:
                if not os.path.exists(image_path):
                    return None
                image = Image.open(image_path)
                base_name = os.path.splitext(os.path.basename(image_path))[0] + '_thumb.jpg'
            
            output_path = os.path.join(self.cache_dir, base_name)
            
            # 创建缩略图
            image.thumbnail(size, Image.Resampling.LANCZOS)
            
            # 保存为JPEG
            if image.mode in ('RGBA', 'LA', 'P'):
                background = Image.new('RGB', image.size, (255, 255, 255))
                if image.mode == 'P':
                    image = image.convert('RGBA')
                background.paste(image, mask=image.split()[-1] if image.mode == 'RGBA' else None)
                image = background
            
            image.save(output_path, 'JPEG', quality=85)
            
            return output_path
            
        except Exception as e:
            print(f"创建缩略图失败 {image_path}: {e}")
            return None
    
    def encode_base64(self, image_path: str) -> Optional[str]:
        """将图片编码为Base64
        
        Args:
            image_path: 图片路径
            
        Returns:
            Optional[str]: Base64编码字符串
        """
        try:
            if self._is_url(image_path):
                image_data = self._download_image(image_path)
                if not image_data:
                    return None
                return base64.b64encode(image_data).decode('utf-8')
            else:
                if not os.path.exists(image_path):
                    return None
                with open(image_path, 'rb') as f:
                    return base64.b64encode(f.read()).decode('utf-8')
                    
        except Exception as e:
            print(f"Base64编码失败 {image_path}: {e}")
            return None
    
    def clear_cache(self) -> bool:
        """清理缓存目录
        
        Returns:
            bool: 清理是否成功
        """
        try:
            import shutil
            if os.path.exists(self.cache_dir):
                shutil.rmtree(self.cache_dir)
            os.makedirs(self.cache_dir, exist_ok=True)
            return True
        except Exception as e:
            print(f"清理缓存失败: {e}")
            return False
    
    def _is_url(self, path: str) -> bool:
        """判断是否为URL"""
        try:
            result = urlparse(path)
            return all([result.scheme, result.netloc])
        except:
            return False
    
    def _download_image(self, url: str, timeout: int = 30) -> Optional[bytes]:
        """下载图片"""
        try:
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            request = Request(url, headers=headers)
            
            with urlopen(request, timeout=timeout) as response:
                return response.read()
                
        except Exception as e:
            print(f"下载图片失败 {url}: {e}")
            return None
    
    def _generate_cache_filename(self, url: str, suffix: str = '') -> str:
        """生成缓存文件名"""
        # 使用URL的哈希值作为文件名
        hash_obj = hashlib.md5(url.encode('utf-8'))
        hash_str = hash_obj.hexdigest()[:12]
        
        # 尝试从URL获取扩展名
        parsed = urlparse(url)
        path = parsed.path
        ext = os.path.splitext(path)[1].lower()
        
        if not ext or ext not in self.supported_formats:
            ext = '.jpg'  # 默认扩展名
        
        return f"{hash_str}{suffix}{ext}"
    
    def _process_image_object(self, image: Image.Image, resize: bool = True) -> Image.Image:
        """处理图片对象"""
        # 自动旋转（基于EXIF信息）
        image = ImageOps.exif_transpose(image)
        
        # 调整大小
        if resize and (image.width > self.max_width or image.height > self.max_height):
            image.thumbnail((self.max_width, self.max_height), Image.Resampling.LANCZOS)
        
        return image
    
    def _save_image(self, image: Image.Image, output_path: str, quality: int = 85):
        """保存图片"""
        # 确保目录存在
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # 根据文件扩展名确定格式
        ext = os.path.splitext(output_path)[1].lower()
        
        if ext in ['.jpg', '.jpeg']:
            # JPEG不支持透明度
            if image.mode in ('RGBA', 'LA', 'P'):
                background = Image.new('RGB', image.size, (255, 255, 255))
                if image.mode == 'P':
                    image = image.convert('RGBA')
                background.paste(image, mask=image.split()[-1] if image.mode == 'RGBA' else None)
                image = background
            image.save(output_path, 'JPEG', quality=quality, optimize=True)
        elif ext == '.png':
            image.save(output_path, 'PNG', optimize=True)
        elif ext == '.webp':
            image.save(output_path, 'WebP', quality=quality, optimize=True)
        else:
            # 默认保存为PNG
            image.save(output_path, 'PNG', optimize=True)
    
    def _has_transparency(self, image: Image.Image) -> bool:
        """检查图片是否有透明度"""
        return image.mode in ('RGBA', 'LA') or (image.mode == 'P' and 'transparency' in image.info)
    
    def _calculate_new_size(self, original_size: Tuple[int, int], 
                           width: Optional[int], height: Optional[int], 
                           maintain_aspect: bool) -> Tuple[int, int]:
        """计算新尺寸"""
        orig_width, orig_height = original_size
        
        if not width and not height:
            return original_size
        
        if maintain_aspect:
            if width and height:
                # 两个都指定时，选择较小的缩放比例
                width_ratio = width / orig_width
                height_ratio = height / orig_height
                ratio = min(width_ratio, height_ratio)
                return (int(orig_width * ratio), int(orig_height * ratio))
            elif width:
                ratio = width / orig_width
                return (width, int(orig_height * ratio))
            else:  # height
                ratio = height / orig_height
                return (int(orig_width * ratio), height)
        else:
            new_width = width or orig_width
            new_height = height or orig_height
            return (new_width, new_height)
    
    def get_cache_size(self) -> Dict[str, Any]:
        """获取缓存大小信息
        
        Returns:
            Dict[str, Any]: 缓存信息
        """
        try:
            total_size = 0
            file_count = 0
            
            if os.path.exists(self.cache_dir):
                for root, dirs, files in os.walk(self.cache_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        total_size += os.path.getsize(file_path)
                        file_count += 1
            
            return {
                'total_size_bytes': total_size,
                'total_size_human': self._format_size(total_size),
                'file_count': file_count,
                'cache_dir': self.cache_dir
            }
            
        except Exception as e:
            print(f"获取缓存信息失败: {e}")
            return {}
    
    def _format_size(self, size_bytes: int) -> str:
        """格式化文件大小"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.1f} TB"