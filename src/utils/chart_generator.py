#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图表生成模块
基于matplotlib生成饼图
"""

import os
import tempfile
from typing import Dict, List, Optional
import matplotlib.pyplot as plt
import matplotlib
from PIL import Image


class ChartGenerator:
    """图表生成器"""
    
    # 默认配色方案
    DEFAULT_COLORS = ['#4A90E2', '#FF8C42', '#808080', '#FFD700', '#1E88E5', '#4CAF50', 
                      '#FF6B6B', '#4ECDC4', '#95E1D3', '#F38181', '#AA96DA', '#FCBAD3']
    
    def __init__(self, output_dir: Optional[str] = None):
        """初始化生成器
        
        Args:
            output_dir: 图片输出目录，如果不提供则使用临时目录
        """
        self.output_dir = output_dir or tempfile.gettempdir()
        self._setup_fonts()
    
    def _setup_fonts(self):
        """设置中文字体"""
        plt.rcParams['font.sans-serif'] = [
            'WenQuanYi Micro Hei', 
            'SimHei', 
            'DejaVu Sans', 
            'Arial Unicode MS', 
            'sans-serif'
        ]
        plt.rcParams['axes.unicode_minus'] = False
    
    def generate_pie_chart(
        self, 
        title: str, 
        data: Dict[str, float], 
        colors: Optional[List[str]] = None,
        width_cm: float = 14.0,
        dpi: int = 300
    ) -> str:
        """生成饼图
        
        Args:
            title: 图表标题
            data: 数据字典，格式：{"标签": 数值}
            colors: 颜色列表，如果不提供则使用默认配色
            width_cm: 图片宽度（厘米）
            dpi: 图片分辨率
            
        Returns:
            生成的图片文件路径
        """
        if not data or len(data) == 0:
            raise ValueError("数据不能为空")
        
        # 准备数据
        labels = list(data.keys())
        sizes = list(data.values())
        
        # 确保数值为正数
        sizes = [max(0, float(s)) for s in sizes]
        
        # 如果所有数值为0，返回错误
        if sum(sizes) == 0:
            raise ValueError("所有数据值不能为0")
        
        # 使用提供的颜色或默认颜色
        chart_colors = colors or self.DEFAULT_COLORS
        # 如果数据项多于颜色，循环使用颜色
        if len(sizes) > len(chart_colors):
            chart_colors = (chart_colors * ((len(sizes) // len(chart_colors)) + 1))[:len(sizes)]
        else:
            chart_colors = chart_colors[:len(sizes)]
        
        # 创建图形，高度根据数据项数量自适应（正常大小）
        height_cm = max(8.0, min(12.0, 6.0 + len(sizes) * 0.5))
        fig, ax = plt.subplots(figsize=(width_cm/2.54, height_cm/2.54), dpi=dpi)
        
        # 绘制饼图
        wedges, texts, autotexts = ax.pie(
            sizes,
            labels=None,  # 不显示标签在图上
            colors=chart_colors,
            autopct='%1.0f%%',  # 显示百分比，不带小数
            startangle=90,
            textprops={'fontsize': 10, 'weight': 'bold', 'color': 'white'}  # 百分比文字样式
        )
        
        # 设置百分比文字背景为黑色矩形
        for autotext in autotexts:
            autotext.set_bbox(dict(
                boxstyle='round,pad=0.3', 
                facecolor='black', 
                edgecolor='none', 
                alpha=0.7
            ))
        
        # 添加图例在右侧
        ax.legend(
            wedges, 
            labels, 
            loc="center left", 
            bbox_to_anchor=(1, 0, 0.5, 1), 
            fontsize=10
        )
        
        # 设置标题
        plt.title(title, fontsize=14, fontweight='bold', pad=20)
        
        # 确保饼图是圆的
        ax.set_aspect('equal')
        
        # 调整布局，为图例留出空间
        plt.tight_layout()
        
        # 生成临时文件路径
        import uuid
        filename = f"chart_{uuid.uuid4().hex[:8]}.png"
        filepath = os.path.join(self.output_dir, filename)
        
        # 保存图片
        print(f"正在保存图片到: {filepath}, DPI: {dpi}")
        plt.savefig(
            filepath,
            dpi=dpi,
            bbox_inches='tight',
            facecolor='white',
            transparent=False
        )
        
        # 关闭图形以释放内存
        plt.close(fig)
        
        # 优化图片文件大小（压缩PNG）
        try:
            img = Image.open(filepath)
            # 如果图片太大，进行压缩
            original_size = os.path.getsize(filepath)
            if original_size > 500 * 1024:  # 如果大于500KB
                print(f"图片文件较大 ({original_size / 1024:.2f} KB)，进行压缩...")
                # 使用PIL重新保存以压缩
                img.save(filepath, 'PNG', optimize=True, compress_level=6)
                new_size = os.path.getsize(filepath)
                print(f"压缩完成: {original_size / 1024:.2f} KB -> {new_size / 1024:.2f} KB")
            else:
                print(f"图片文件大小: {original_size / 1024:.2f} KB")
            img.close()
        except Exception as e:
            print(f"图片压缩失败（继续使用原图）: {e}")
        
        print(f"图片保存完成: {filepath}")
        return filepath
    
    def cleanup(self, filepath: str):
        """清理临时图片文件
        
        Args:
            filepath: 要删除的图片文件路径
        """
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
        except Exception as e:
            print(f"清理图片文件失败 {filepath}: {e}")

