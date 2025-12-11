#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图表生成模块
基于matplotlib生成饼图
"""

import os
import tempfile
from typing import Dict, List, Optional
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
        data: Dict[str, float], 
        colors: Optional[List[str]] = None,
        width_cm: float = 14.0,
        dpi: int = 300
    ) -> str:
        """生成柱状图
        
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
        values = list(data.values())
        
        # 确保数值为正数
        values = [max(0, float(v)) for v in values]
        
        # 如果所有数值为0，返回错误
        if sum(values) == 0:
            raise ValueError("所有数据值不能为0")
        
        # 使用提供的颜色或默认颜色
        chart_colors = colors or self.DEFAULT_COLORS
        
        # 创建图形，高度根据数据项数量自适应，设置最低高度为12cm以确保美观
        height_cm = max(12.0, min(16.0, 8.0 + len(values) * 0.8))
        fig, ax = plt.subplots(figsize=(width_cm/2.54, height_cm/2.54), dpi=dpi)
        
        # 绘制柱状图
        bars = ax.bar(
            range(len(labels)),
            values,
            color=[chart_colors[i % len(chart_colors)] for i in range(len(labels))]
        )
        
        # 设置标题和字体（使用字体文件路径确保中文显示）
        from matplotlib.font_manager import FontProperties
        if self._font_file_path:
            font_prop = FontProperties(fname=self._font_file_path)
            plt.title(title, fontsize=14, fontweight='bold', pad=20, fontproperties=font_prop)
            # 设置y轴标签字体
            ax.set_ylabel('数值', fontproperties=font_prop, fontsize=12)
        else:
            plt.title(title, fontsize=14, fontweight='bold', pad=20)
            ax.set_ylabel('数值', fontsize=12)
        
        # 设置x轴标签（确保中文正常显示）
        ax.set_xticks(range(len(labels)))
        if self._font_file_path:
            # 为每个标签设置字体属性，确保中文正常显示
            # 使用较小的旋转角度（30度）和更好的对齐方式，避免标签被截断
            ax.set_xticklabels(labels, rotation=30, ha='right', fontproperties=font_prop, fontsize=10)
        else:
            ax.set_xticklabels(labels, rotation=30, ha='right', fontsize=10)
        
        # 计算y轴上限，为数值标签留出空间（增加最大值15%的空间）
        max_value = max(values) if values else 1
        y_max = max_value * 1.15
        ax.set_ylim(0, y_max)
        
        # 在柱状图上显示数值（位置在柱子顶部上方，避免被边框遮挡）
        for i, (bar, value) in enumerate(zip(bars, values)):
            height = bar.get_height()
            # 在柱子高度基础上增加3-5%的偏移，确保标签在柱子顶部上方且不被边框遮挡
            label_y = height + max_value * 0.03
            if self._font_file_path:
                ax.text(
                    bar.get_x() + bar.get_width() / 2., 
                    label_y,
                    f'{value:.0f}',
                    ha='center', 
                    va='bottom',
                    fontsize=10,
                    fontweight='bold',
                    fontproperties=font_prop
                )
            else:
                ax.text(
                    bar.get_x() + bar.get_width() / 2., 
                    label_y,
                    f'{value:.0f}',
                    ha='center', 
                    va='bottom',
                    fontsize=10,
                    fontweight='bold'
                )
        
        # 设置网格
        ax.grid(axis='y', alpha=0.3, linestyle='--')
        
        # 调整布局，为旋转的x轴标签留出足够的底部空间
        # 根据标签数量和长度动态调整底部边距
        max_label_length = max([len(str(label)) for label in labels]) if labels else 0
        num_labels = len(labels) if labels else 0
        # 动态计算底部边距：基础0.35 + 标签长度影响 + 标签数量影响
        # 旋转30度的中文标签需要足够的底部空间，确保长标签完全可见
        bottom_margin = max(0.35, 0.30 + max_label_length * 0.03 + num_labels * 0.025)
        # 先使用 subplots_adjust 设置精确的边距
        plt.subplots_adjust(
            bottom=bottom_margin,  # 底部边距
            top=0.90,              # 顶部边距
            left=0.12,             # 左边距（为y轴标签留空间）
            right=0.95             # 右边距
        )
        # 然后再使用 tight_layout 进行微调
        plt.tight_layout(rect=[0, bottom_margin, 1, 0.90])
        
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
    
    def generate_line_chart(
        self, 
        title: str, 
        data: Dict[str, float], 
        colors: Optional[List[str]] = None,
        width_cm: float = 14.0,
        dpi: int = 300
    ) -> str:
        """生成折线图
        
        Args:
            title: 图表标题
            data: 数据字典，格式：{"标签": 数值}，标签通常为时间或顺序值
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

