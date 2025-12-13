#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图表生成模块
基于matplotlib生成饼图
"""

import os
import tempfile
from typing import Dict, List, Optional, Union
import numpy as np
import matplotlib
# 在 Docker 环境中使用无界面后端
matplotlib.use('Agg')
import matplotlib.pyplot as plt
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
        self._font_file_path = None  # 保存字体文件路径
        self._setup_fonts()
    
    def _setup_fonts(self):
        """设置中文字体，支持 Docker 环境，优先使用项目中的字体文件"""
        import matplotlib.font_manager as fm
        from urllib.request import urlretrieve
        from pathlib import Path
        
        # 获取系统中所有可用字体
        available_fonts = [f.name for f in fm.fontManager.ttflist]
        
        # 中文字体优先级列表（按可用性排序，排除不支持中文的字体）
        chinese_fonts = [
            'WenQuanYi Micro Hei',      # 文泉驿微米黑（Linux常用）
            'WenQuanYi Zen Hei',        # 文泉驿正黑
            'Noto Sans CJK SC',         # Noto Sans 中文字体
            'Noto Sans SC',             # Noto Sans 简体中文
            'Source Han Sans SC',       # 思源黑体
            'SimHei',                   # 黑体（Windows）
            'Microsoft YaHei',          # 微软雅黑（Windows）
            'SimSun',                   # 宋体（Windows）
            'STHeiti',                  # 华文黑体（macOS）
            'STSong',                   # 华文宋体（macOS）
            'Arial Unicode MS',         # Arial Unicode（跨平台）
        ]
        
        # 不支持中文的字体（需要排除）
        non_chinese_fonts = ['DejaVu Sans', 'Arial', 'Helvetica', 'Times New Roman']
        
        selected_font = None
        
        # 第一步：优先使用项目中的字体文件（最重要！）
        print("步骤1: 检查项目中的字体文件...")
        current_file = Path(__file__).resolve()
        # 从 src/utils/chart_generator.py 向上找到项目根目录
        project_root = current_file.parent.parent.parent
        # 字体文件路径（支持多种格式和文件名）
        fonts_dir = project_root / 'src' / 'assets' / 'fonts'
        font_file = None
        # 尝试多种可能的字体文件名和格式
        possible_font_names = [
            'noto-sans-sc-regular.otf',  # OTF 格式
            'NotoSansSC-Regular.otf',
            'NotoSansSC-Regular.ttf',    # TTF 格式
            'noto-sans-sc-regular.ttf',
        ]
        for font_name in possible_font_names:
            candidate = fonts_dir / font_name
            if candidate.exists():
                font_file = candidate
                print(f"找到字体文件: {font_file}")
                break
        
        # 如果项目中有字体文件，直接使用
        if font_file and font_file.exists():
            print(f"找到项目中的字体文件: {font_file}")
            try:
                # 添加字体文件到 matplotlib
                fm.fontManager.addfont(str(font_file))
                # 重新构建字体缓存（使用安全的方法）
                try:
                    # 尝试清除缓存并重新加载
                    import matplotlib
                    matplotlib.font_manager._rebuild()
                except AttributeError:
                    # 如果 _rebuild 不存在，尝试其他方法
                    try:
                        fm.fontManager.__init__()
                    except:
                        pass
                
                # 获取字体文件的实际字体名称
                try:
                    from matplotlib.font_manager import FontProperties
                    font_prop = FontProperties(fname=str(font_file))
                    actual_font_name = font_prop.get_name()
                    print(f"字体文件实际名称: {actual_font_name}")
                    selected_font = actual_font_name
                except:
                    # 如果获取失败，尝试常见的名称
                    possible_names = ['Noto Sans SC', 'NotoSansSC-Regular', 'Noto Sans CJK SC']
                    for name in possible_names:
                        # 检查字体是否在系统中
                        if name in [f.name for f in fm.fontManager.ttflist]:
                            selected_font = name
                            break
                    if not selected_font:
                        selected_font = 'Noto Sans SC'
                
                # 保存字体文件路径，用于后续直接使用
                self._font_file_path = str(font_file)
                print(f"✅ 成功使用项目中的字体: {selected_font} (文件: {font_file.name})")
            except Exception as e:
                print(f"❌ 注册项目字体失败: {e}")
                import traceback
                traceback.print_exc()
                selected_font = None
                self._font_file_path = None
        else:
            print(f"❌ 项目字体文件不存在: {font_file}")
        
        # 第二步：如果项目字体加载失败，再检查系统字体
        if selected_font is None:
            print("步骤2: 检查系统中的中文字体...")
            # 查找第一个可用的中文字体（排除不支持中文的字体）
            for font_name in chinese_fonts:
                if font_name in available_fonts and font_name not in non_chinese_fonts:
                    selected_font = font_name
                    print(f"✅ 找到系统可用中文字体: {font_name}")
                    break
            
            # 如果项目中没有字体文件或注册失败，尝试下载
            if selected_font is None:
                try:
                    # 创建字体缓存目录
                    font_cache_dir = Path.home() / '.matplotlib' / 'fonts'
                    font_cache_dir.mkdir(parents=True, exist_ok=True)
                    
                    # 尝试下载 Noto Sans SC 字体（Google 提供的免费中文字体）
                    downloaded_font_file = font_cache_dir / 'NotoSansSC-Regular.ttf'
                    if not downloaded_font_file.exists():
                        print("正在从网络下载 Noto Sans SC 字体...")
                        try:
                            # 使用 Google Fonts 的 CDN 下载字体
                            font_url = 'https://github.com/google/fonts/raw/main/ofl/notosanssc/NotoSansSC%5Bwdth%2Cwght%5D.ttf'
                            urlretrieve(font_url, downloaded_font_file)
                            print(f"字体下载成功: {downloaded_font_file}")
                        except Exception as e:
                            print(f"字体下载失败: {e}，尝试备用方案...")
                            # 备用：使用 GitHub 的字体文件
                            try:
                                font_url = 'https://raw.githubusercontent.com/googlefonts/noto-cjk/main/Sans/Variable/TTF/Subset/NotoSansCJKsc-VF.ttf'
                                urlretrieve(font_url, downloaded_font_file)
                                print(f"使用备用字体源下载成功: {downloaded_font_file}")
                            except Exception as e2:
                                print(f"备用字体下载也失败: {e2}")
                                downloaded_font_file = None
                    else:
                        print(f"使用已缓存的字体文件: {downloaded_font_file}")
                    
                    # 如果字体文件存在，注册并使用它
                    if downloaded_font_file and downloaded_font_file.exists():
                        try:
                            # 添加字体文件到 matplotlib
                            fm.fontManager.addfont(str(downloaded_font_file))
                            # 重新构建字体缓存（使用安全的方法）
                            try:
                                # 尝试清除缓存并重新加载
                                import matplotlib
                                matplotlib.font_manager._rebuild()
                            except AttributeError:
                                # 如果 _rebuild 不存在，尝试其他方法
                                try:
                                    fm.fontManager.__init__()
                                except:
                                    pass
                            # 使用下载的字体
                            selected_font = 'Noto Sans SC'
                            print(f"成功使用下载的字体: {selected_font}")
                        except Exception as e:
                            print(f"注册下载字体失败: {e}")
                            selected_font = None
                except Exception as e:
                    print(f"字体下载过程出错: {e}")
                    selected_font = None
        
        # 如果仍然没有找到中文字体，使用备用方案
        if selected_font is None:
            print("警告: 无法获取中文字体，使用备用字体（中文可能显示为方块）")
            # 尝试查找任何包含中文的字体
            try:
                noto_fonts = [f for f in available_fonts if 'noto' in f.lower() or 'cjk' in f.lower()]
                if noto_fonts:
                    selected_font = noto_fonts[0]
                    print(f"使用 Noto 字体: {selected_font}")
                else:
                    selected_font = 'DejaVu Sans'
                    print(f"使用备用字体 {selected_font}，中文将显示为方块")
            except Exception as e:
                print(f"字体检测失败: {e}")
                selected_font = 'DejaVu Sans'
        
        # 设置字体（确保不使用不支持中文的字体作为主字体）
        if selected_font and selected_font not in non_chinese_fonts:
            font_list = [selected_font] + [f for f in chinese_fonts if f != selected_font] + ['sans-serif']
        else:
            # 如果主字体不支持中文，至少尝试使用项目字体
            font_list = ['Noto Sans SC'] + chinese_fonts + ['sans-serif']
            print(f"警告: 主字体可能不支持中文，使用字体列表: {font_list[:3]}...")
        
        plt.rcParams['font.sans-serif'] = font_list
        plt.rcParams['axes.unicode_minus'] = False
        
        # 清除字体缓存，确保使用最新字体
        try:
            # 尝试重新构建字体缓存
            import matplotlib
            matplotlib.font_manager._rebuild()
        except (AttributeError, Exception):
            # 如果重建失败，忽略错误（字体设置仍然有效）
            pass
    
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
        
        # 计算百分比
        total = sum(sizes)
        percentages = [s / total * 100 for s in sizes]
        
        # 检查是否有小于8%的分块
        has_small_slices = any(pct < 8 for pct in percentages)
        
        if not has_small_slices:
            # 如果所有分块都大于等于8%，使用均匀分布（autopct自动标注）
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
        else:
            # 如果有小于8%的分块，使用错开标注
            wedges, texts, _ = ax.pie(
                sizes,
                labels=None,  # 不显示标签在图上
                colors=chart_colors,
                autopct='',  # 不自动显示百分比
                startangle=90
            )
            
            # 手动添加百分比标注，根据切片大小决定位置
            from matplotlib.patches import ConnectionPatch
            
            # 获取每个楔形的角度中心
            for i, (wedge, pct) in enumerate(zip(wedges, percentages)):
                # 计算楔形的角度中心（弧度）
                theta1, theta2 = wedge.theta1, wedge.theta2
                theta_center = np.deg2rad((theta1 + theta2) / 2)
                
                # 计算楔形的中心点坐标
                # 根据切片大小决定标注位置
                if pct >= 8:
                    # 大切片（>=8%）：将百分比放在内部，距离圆心较近
                    # 使用不同的距离避免重叠
                    distance = 0.6 + (i % 3) * 0.1  # 错开距离：0.6, 0.7, 0.8
                    x = distance * np.cos(theta_center)
                    y = distance * np.sin(theta_center)
                    ha = 'center'
                    va = 'center'
                    use_connection = False
                else:
                    # 小切片（<8%）：将百分比放在外部，使用引线连接
                    distance = 1.15 + (i % 2) * 0.1  # 错开距离：1.15, 1.25
                    x = distance * np.cos(theta_center)
                    y = distance * np.sin(theta_center)
                    # 根据角度决定水平对齐方式
                    if abs(np.cos(theta_center)) > abs(np.sin(theta_center)):
                        ha = 'left' if np.cos(theta_center) > 0 else 'right'
                        va = 'center'
                    else:
                        ha = 'center'
                        va = 'bottom' if np.sin(theta_center) > 0 else 'top'
                    use_connection = True
                    
                    # 添加引线（从小切片边缘到标注位置）
                    # 计算切片边缘点
                    edge_x = 1.05 * np.cos(theta_center)
                    edge_y = 1.05 * np.sin(theta_center)
                    # 创建连接线
                    con = ConnectionPatch(
                        (edge_x, edge_y), (x, y), "data", "data",
                        arrowstyle="-", shrinkA=0, shrinkB=0,
                        mutation_scale=20, fc="gray", alpha=0.5, linestyle='--'
                    )
                    ax.add_patch(con)
                
                # 添加百分比文本
                from matplotlib.font_manager import FontProperties
                # 根据是否为外部标注决定文字颜色
                text_color = 'white' if not use_connection else 'black'
                
                if self._font_file_path:
                    font_prop = FontProperties(fname=self._font_file_path)
                    text = ax.text(
                        x, y, f'{pct:.0f}%',
                        ha=ha, va=va,
                        fontsize=10, weight='bold', color=text_color,
                        fontproperties=font_prop
                    )
                else:
                    text = ax.text(
                        x, y, f'{pct:.0f}%',
                        ha=ha, va=va,
                        fontsize=10, weight='bold', color=text_color
                    )
                
                # 只为内部标注添加黑色背景，外部标注不添加背景
                if not use_connection:
                    # 内部标注（>=8%）：白字黑底
                    text.set_bbox(dict(
                        boxstyle='round,pad=0.3',
                        facecolor='black',
                        edgecolor='none',
                        alpha=0.7
                    ))
                # 外部标注（<8%）：黑字无背景
        
        # 添加图例在右侧（使用字体文件路径确保中文显示）
        from matplotlib.font_manager import FontProperties
        if self._font_file_path:
            font_prop = FontProperties(fname=self._font_file_path)
            ax.legend(
                wedges, 
                labels, 
                loc="center left", 
                bbox_to_anchor=(1, 0, 0.5, 1), 
                fontsize=10,
                prop=font_prop
            )
            # 设置标题（使用字体文件路径）
            plt.title(title, fontsize=14, fontweight='bold', pad=20, fontproperties=font_prop)
        else:
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
    
    def generate_bar_chart(
        self, 
        title: str, 
        data: Union[Dict[str, float], Dict[str, Dict[str, float]]], 
        colors: Optional[List[str]] = None,
        width_cm: float = 14.0,
        dpi: int = 300
    ) -> str:
        """生成柱状图（支持单一数据系列和分组数据系列）
        
        Args:
            title: 图表标题
            data: 数据字典，支持两种格式：
                  - 单一系列：{"标签": 数值}，如 {"4月": 100, "5月": 200}
                  - 分组系列：{"系列名": {"标签": 数值}}，如 {"4级告警": {"4月": 1519, "5月": 1616}, "5级告警": {"4月": 73, "5月": 164}}
            colors: 颜色列表，如果不提供则使用默认配色
            width_cm: 图片宽度（厘米）
            dpi: 图片分辨率
            
        Returns:
            生成的图片文件路径
        """
        if not data or len(data) == 0:
            raise ValueError("数据不能为空")
        
        # 检测数据格式：判断是否为分组数据
        is_grouped = False
        first_value = next(iter(data.values()))
        if isinstance(first_value, dict):
            is_grouped = True
        
        if is_grouped:
            return self._generate_grouped_bar_chart(title, data, colors, width_cm, dpi)
        else:
            return self._generate_single_bar_chart(title, data, colors, width_cm, dpi)
    
    def _generate_single_bar_chart(
        self,
        title: str,
        data: Dict[str, float],
        colors: Optional[List[str]],
        width_cm: float,
        dpi: int
    ) -> str:
        """生成单一数据系列的柱状图"""
        # 准备数据
        labels = list(data.keys())
        values = [float(v) for v in data.values()]
        
        # 使用提供的颜色或默认颜色
        chart_colors = colors or self.DEFAULT_COLORS
        
        # 创建图形
        height_cm = 10.0
        fig, ax = plt.subplots(figsize=(width_cm/2.54, height_cm/2.54), dpi=dpi)
        
        # 绘制柱状图
        bars = ax.bar(
            range(len(labels)),
            values,
            color=[chart_colors[i % len(chart_colors)] for i in range(len(labels))]
        )
        
        # 设置字体（使用与饼图相同的字体设置）
        from matplotlib.font_manager import FontProperties
        if self._font_file_path:
            font_prop = FontProperties(fname=self._font_file_path)
            # 设置标题和y轴标签
            plt.title(title, fontsize=14, fontweight='bold', fontproperties=font_prop)
            ax.set_ylabel('数值', fontproperties=font_prop, fontsize=12)
            
            # 设置x轴标签
            ax.set_xticks(range(len(labels)))
            ax.set_xticklabels(labels, rotation=30, ha='right', fontproperties=font_prop, fontsize=10)
        else:
            # 设置标题和y轴标签
            plt.title(title, fontsize=14, fontweight='bold')
            ax.set_ylabel('数值', fontsize=12)
            
            # 设置x轴标签
            ax.set_xticks(range(len(labels)))
            ax.set_xticklabels(labels, rotation=30, ha='right', fontsize=10)
        
        # 设置y轴范围
        max_value = max(values) if values else 1
        ax.set_ylim(0, max_value * 1.2)
        
        # 在柱状图上显示数值
        for bar, value in zip(bars, values):
            height = bar.get_height()
            if self._font_file_path:
                ax.text(
                    bar.get_x() + bar.get_width() / 2., 
                    height,
                    f'{value:.0f}',
                    ha='center', 
                    va='bottom',
                    fontsize=10,
                    fontproperties=font_prop
                )
            else:
                ax.text(
                    bar.get_x() + bar.get_width() / 2., 
                    height,
                    f'{value:.0f}',
                    ha='center', 
                    va='bottom',
                    fontsize=10
                )
        
        # 设置网格
        ax.grid(axis='y', alpha=0.3)
        
        # 调整布局
        plt.tight_layout()
        
        # 保存图片
        import uuid
        filename = f"chart_{uuid.uuid4().hex[:8]}.png"
        filepath = os.path.join(self.output_dir, filename)
        plt.savefig(filepath, dpi=dpi, bbox_inches='tight', facecolor='white')
        plt.close(fig)
        
        return filepath
    
    def _generate_grouped_bar_chart(
        self,
        title: str,
        data: Dict[str, Dict[str, float]],
        colors: Optional[List[str]],
        width_cm: float,
        dpi: int
    ) -> str:
        """生成分组柱状图（多个数据系列）"""
        # 准备数据
        series_names = list(data.keys())  # 系列名称，如 ["4级告警", "5级告警", "总量告警"]
        
        # 获取所有标签（x轴标签，如月份）
        all_labels = set()
        for series_data in data.values():
            all_labels.update(series_data.keys())
        labels = sorted(list(all_labels))  # 排序以确保顺序一致
        
        # 使用提供的颜色或默认颜色
        chart_colors = colors or self.DEFAULT_COLORS
        
        # 创建图形，根据系列数量和标签数量调整宽度和高度
        num_series = len(series_names)
        num_labels = len(labels)
        # 增加宽度以适应更多标签和更宽的柱子
        adjusted_width_cm = max(width_cm, width_cm * (1 + num_labels * 0.1))
        height_cm = max(10.0, 8.0 + num_series * 0.3)
        fig, ax = plt.subplots(figsize=(adjusted_width_cm/2.54, height_cm/2.54), dpi=dpi)
        
        # 计算柱状图的位置和宽度
        # 增加组间距，减少柱子拥挤
        x = np.arange(len(labels))  # x轴位置
        # 增加每个柱子的宽度，减少拥挤感
        bar_width = 0.25  # 固定每个柱子的宽度，不再除以系列数
        # 计算组内间距（同一组内柱子之间的间距）
        group_spacing = bar_width * 0.3  # 组内间距为柱子宽度的30%
        # 计算总组宽度
        total_group_width = num_series * bar_width + (num_series - 1) * group_spacing
        # 计算偏移量，使柱状图居中
        offset = total_group_width / 2 - bar_width / 2
        
        # 存储每个系列的颜色，确保图例颜色一致
        series_colors = []
        
        # 绘制每个系列的柱状图
        bars_list = []
        for i, series_name in enumerate(series_names):
            series_data = data[series_name]
            values = [float(series_data.get(label, 0)) for label in labels]
            
            # 选择颜色
            series_color = chart_colors[i % len(chart_colors)]
            series_colors.append(series_color)
            
            # 计算每个系列的位置
            x_pos = x - offset + i * (bar_width + group_spacing)
            
            # 绘制柱状图
            bars = ax.bar(
                x_pos,
                values,
                bar_width,
                label=series_name,
                color=series_color
            )
            bars_list.append(bars)
            
            # 在柱状图上显示数值
            from matplotlib.font_manager import FontProperties
            font_prop = FontProperties(fname=self._font_file_path) if self._font_file_path else None
            
            for bar, value in zip(bars, values):
                if value > 0:  # 只显示非零值
                    height = bar.get_height()
                    if font_prop:
                        ax.text(
                            bar.get_x() + bar.get_width() / 2.,
                            height,
                            f'{value:.0f}',
                            ha='center',
                            va='bottom',
                            fontsize=8,
                            fontproperties=font_prop
                        )
                    else:
                        ax.text(
                            bar.get_x() + bar.get_width() / 2.,
                            height,
                            f'{value:.0f}',
                            ha='center',
                            va='bottom',
                            fontsize=8
                        )
        
        # 设置字体
        from matplotlib.font_manager import FontProperties
        if self._font_file_path:
            font_prop = FontProperties(fname=self._font_file_path)
            # 设置标题和y轴标签
            plt.title(title, fontsize=14, fontweight='bold', fontproperties=font_prop)
            ax.set_ylabel('数值', fontproperties=font_prop, fontsize=12)
            
            # 设置x轴标签
            ax.set_xticks(x)
            ax.set_xticklabels(labels, rotation=30, ha='right', fontproperties=font_prop, fontsize=10)
            
            # 设置图例 - 移到图表外部右侧，避免与柱子重叠
            # 使用 handles 和 labels 确保颜色一致
            from matplotlib.patches import Rectangle
            handles = [Rectangle((0,0),1,1, facecolor=color, edgecolor='none') 
                      for color in series_colors]
            ax.legend(
                handles=handles,
                labels=series_names,
                loc='center left',
                bbox_to_anchor=(1.02, 0.5),
                fontsize=10,
                prop=font_prop,
                frameon=True,
                fancybox=True,
                shadow=True
            )
        else:
            # 设置标题和y轴标签
            plt.title(title, fontsize=14, fontweight='bold')
            ax.set_ylabel('数值', fontsize=12)
            
            # 设置x轴标签
            ax.set_xticks(x)
            ax.set_xticklabels(labels, rotation=30, ha='right', fontsize=10)
            
            # 设置图例 - 移到图表外部右侧，避免与柱子重叠
            from matplotlib.patches import Rectangle
            handles = [Rectangle((0,0),1,1, facecolor=color, edgecolor='none') 
                      for color in series_colors]
            ax.legend(
                handles=handles,
                labels=series_names,
                loc='center left',
                bbox_to_anchor=(1.02, 0.5),
                fontsize=10,
                frameon=True,
                fancybox=True,
                shadow=True
            )
        
        # 设置y轴范围
        all_values = []
        for series_data in data.values():
            all_values.extend([float(v) for v in series_data.values()])
        max_value = max(all_values) if all_values else 1
        ax.set_ylim(0, max_value * 1.2)
        
        # 设置网格
        ax.grid(axis='y', alpha=0.3)
        
        # 调整布局，为右侧图例留出空间
        plt.tight_layout(rect=[0, 0, 0.85, 1])  # 左侧0，底部0，右侧0.85（留出15%给图例），顶部1
        
        # 保存图片
        import uuid
        filename = f"chart_{uuid.uuid4().hex[:8]}.png"
        filepath = os.path.join(self.output_dir, filename)
        plt.savefig(filepath, dpi=dpi, bbox_inches='tight', facecolor='white')
        plt.close(fig)
        
        return filepath
    
    def generate_line_chart(
        self, 
        title: str, 
        data: Union[Dict[str, float], Dict[str, Dict[str, float]]], 
        colors: Optional[List[str]] = None,
        width_cm: float = 14.0,
        dpi: int = 300
    ) -> str:
        """生成折线图（支持单一数据系列和多个数据系列）
        
        Args:
            title: 图表标题
            data: 数据字典，支持两种格式：
                  - 单一系列：{"标签": 数值}，如 {"4月": 100, "5月": 200}
                  - 多条系列：{"系列名": {"标签": 数值}}，如 {"主要告警": {"4月": 1565, "5月": 1762}, "硬件监控告警": {"4月": 27, "5月": 18}}
            colors: 颜色列表，如果不提供则使用默认配色
            width_cm: 图片宽度（厘米）
            dpi: 图片分辨率
            
        Returns:
            生成的图片文件路径
        """
        if not data or len(data) == 0:
            raise ValueError("数据不能为空")
        
        # 检测数据格式：判断是否为多条数据
        is_multi_line = False
        first_value = next(iter(data.values()))
        if isinstance(first_value, dict):
            is_multi_line = True
        
        if is_multi_line:
            return self._generate_multi_line_chart(title, data, colors, width_cm, dpi)
        else:
            return self._generate_single_line_chart(title, data, colors, width_cm, dpi)
    
    def _generate_single_line_chart(
        self,
        title: str,
        data: Dict[str, float],
        colors: Optional[List[str]],
        width_cm: float,
        dpi: int
    ) -> str:
        """生成单一数据系列的折线图"""
        # 准备数据
        labels = list(data.keys())
        values = list(data.values())
        
        # 确保数值为数字
        values = [float(v) for v in values]
        
        # 使用提供的颜色或默认颜色
        chart_colors = colors or self.DEFAULT_COLORS
        line_color = chart_colors[0]  # 折线图通常使用单一颜色
        
        # 创建图形，高度固定
        height_cm = 10.0
        fig, ax = plt.subplots(figsize=(width_cm/2.54, height_cm/2.54), dpi=dpi)
        
        # 设置标题和字体（使用字体文件路径确保中文显示）
        from matplotlib.font_manager import FontProperties
        if self._font_file_path:
            font_prop = FontProperties(fname=self._font_file_path)
        else:
            font_prop = None
        
        # 绘制折线图
        x_positions = range(len(labels))
        line = ax.plot(
            x_positions,
            values,
            marker='o',  # 在数据点处显示圆点
            linestyle='-',  # 实线
            linewidth=2,
            markersize=6,
            color=line_color,
            label=title
        )
        
        # 设置x轴标签（确保中文正常显示）
        ax.set_xticks(x_positions)
        if font_prop:
            # 为每个标签设置字体属性，确保中文正常显示
            ax.set_xticklabels(labels, rotation=45, ha='right', fontproperties=font_prop, fontsize=10)
        else:
            ax.set_xticklabels(labels, rotation=45, ha='right', fontsize=10)
        
        # 计算y轴范围，为数值标签留出空间
        min_value = min(values) if values else 0
        max_value = max(values) if values else 1
        value_range = max_value - min_value if max_value != min_value else max_value or 1
        # 为y轴留出上下边距
        y_min = min_value - value_range * 0.1
        y_max = max_value + value_range * 0.15
        ax.set_ylim(y_min, y_max)
        
        # 在数据点上显示数值（确保中文字体一致性）
        for i, (x, y, value) in enumerate(zip(x_positions, values, values)):
            # 在数据点上方显示数值
            label_y = y + value_range * 0.03
            if font_prop:
                ax.text(
                    x,
                    label_y,
                    f'{value:.0f}' if value == int(value) else f'{value:.1f}',
                    ha='center',
                    va='bottom',
                    fontsize=9,
                    fontweight='bold',
                    fontproperties=font_prop
                )
            else:
                ax.text(
                    x,
                    label_y,
                    f'{value:.0f}' if value == int(value) else f'{value:.1f}',
                    ha='center',
                    va='bottom',
                    fontsize=9,
                    fontweight='bold'
                )
        
        # 设置标题和轴标签（使用字体文件路径确保中文显示）
        if font_prop:
            plt.title(title, fontsize=14, fontweight='bold', pad=20, fontproperties=font_prop)
            # 设置y轴标签字体
            ax.set_ylabel('数值', fontproperties=font_prop, fontsize=12)
            ax.set_xlabel('', fontproperties=font_prop, fontsize=12)
        else:
            plt.title(title, fontsize=14, fontweight='bold', pad=20)
            ax.set_ylabel('数值', fontsize=12)
            ax.set_xlabel('', fontsize=12)
        
        # 设置网格
        ax.grid(axis='y', alpha=0.3, linestyle='--')
        ax.grid(axis='x', alpha=0.2, linestyle='--')
        
        # 调整布局，为旋转的x轴标签留出足够的底部空间
        max_label_length = max([len(str(label)) for label in labels]) if labels else 0
        num_labels = len(labels) if labels else 0
        # 动态计算底部边距：基础0.15 + 标签长度影响 + 标签数量影响
        bottom_margin = max(0.2, 0.15 + max_label_length * 0.015 + num_labels * 0.01)
        plt.tight_layout(rect=[0, bottom_margin, 1, 0.95])  # 增加底部边距，避免标签被截断
        
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
    
    def _generate_multi_line_chart(
        self,
        title: str,
        data: Dict[str, Dict[str, float]],
        colors: Optional[List[str]],
        width_cm: float,
        dpi: int
    ) -> str:
        """生成多条折线图（多个数据系列）"""
        from matplotlib.font_manager import FontProperties
        import uuid
        
        # 字号配置
        FONT_SIZE = {'title': 11, 'tick': 5.5, 'legend': 5.5, 'value': 4}
        
        # 创建带字号的字体属性（解决 fontsize/fontproperties 冲突问题）
        def get_font(size: float) -> Optional[FontProperties]:
            if self._font_file_path:
                return FontProperties(fname=self._font_file_path, size=size)
            return None
        
        # 准备数据
        series_names = list(data.keys())
        all_labels = sorted(set(label for series in data.values() for label in series.keys()))
        chart_colors = colors or self.DEFAULT_COLORS
        x_pos = range(len(all_labels))
        
        # 创建图形
        height_cm = max(10.0, 8.0 + len(series_names) * 0.3)
        fig, ax = plt.subplots(figsize=(width_cm/2.54, height_cm/2.54), dpi=dpi)
        
        # 收集所有值用于计算Y轴范围
        all_values = []
        
        # 绘制每条折线
        for i, series_name in enumerate(series_names):
            values = [float(data[series_name].get(label, 0)) for label in all_labels]
            all_values.extend(values)
            color = chart_colors[i % len(chart_colors)]
            
            ax.plot(x_pos, values, marker='o', linestyle='-', linewidth=1.2,
                    markersize=4, color=color, label=series_name)
            
            # 数据点标签
            value_range = (max(all_values) - min(all_values)) if len(all_values) > 1 else 1
            font_value = get_font(FONT_SIZE['value'])
            for x, v in zip(x_pos, values):
                if v > 0:
                    text = f'{v:.0f}' if v == int(v) else f'{v:.1f}'
                    ax.text(x, v + value_range * 0.008, text, ha='center', va='bottom',
                            fontproperties=font_value, fontsize=FONT_SIZE['value'])
        
        # Y轴范围
        if all_values:
            v_min, v_max = min(all_values), max(all_values)
            v_range = v_max - v_min if v_max != v_min else v_max or 1
            ax.set_ylim(v_min - v_range * 0.1, v_max + v_range * 0.15)
        
        # X轴标签
        ax.set_xticks(x_pos)
        font_tick = get_font(FONT_SIZE['tick'])
        ax.set_xticklabels(all_labels, rotation=45, ha='right', 
                          fontproperties=font_tick, fontsize=FONT_SIZE['tick'])
        
        # Y轴刻度字体（使用 tick_params 确保生效）
        ax.tick_params(axis='y', labelsize=FONT_SIZE['tick'])
        if font_tick:
            for label in ax.get_yticklabels():
                label.set_fontproperties(font_tick)
        
        # 标题
        font_title = get_font(FONT_SIZE['title'])
        ax.set_title(title, fontproperties=font_title, fontsize=FONT_SIZE['title'], 
                     fontweight='bold', pad=20)
        
        # Y轴标签
        ax.set_ylabel('数值', fontproperties=font_tick, fontsize=FONT_SIZE['tick'])
        
        # 图例
        font_legend = get_font(FONT_SIZE['legend'])
        ax.legend(loc='center left', bbox_to_anchor=(1.005, 0.5),
                  prop=font_legend, frameon=True, fancybox=False, shadow=False,
                  borderpad=0.3, handlelength=1.2, handletextpad=0.3)
        
        # 网格
        ax.grid(axis='both', alpha=0.3, linestyle='--')
        
        # 保存图片
        filepath = os.path.join(self.output_dir, f"chart_{uuid.uuid4().hex[:8]}.png")
        fig.savefig(filepath, dpi=dpi, bbox_inches='tight', facecolor='white')
        plt.close(fig)
        
        # 压缩大图片
        if os.path.getsize(filepath) > 500 * 1024:
            img = Image.open(filepath)
            img.save(filepath, 'PNG', optimize=True, compress_level=6)
            img.close()
        
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

