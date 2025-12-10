import matplotlib.pyplot as plt
import matplotlib

# 配置中文字体
# 尝试使用系统中可用的中文字体
plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans', 'Arial Unicode MS', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

# 定义饼图数据
labels = ['海量数据处理平台', '绩效考核平台', '老财务管理系统', '审计信息系统', '数据备份管理平台', '移动办公统一门户']
sizes = [50, 10, 8, 6, 18, 8]  # 每个部分的比例

# 定义每个部分的颜色（根据右侧图片调整）
colors = ['#4A90E2', '#FF8C42', '#808080', '#FFD700', '#1E88E5', '#4CAF50']

# 创建图形
fig, ax = plt.subplots(figsize=(10, 7))

# 绘制2D饼图
# 绘制饼图，不显示标签和百分比在图上
wedges, texts, autotexts = ax.pie(
    sizes, 
    labels=None,  # 不显示标签在图上
    colors=colors,
    autopct='%1.0f%%',  # 显示百分比，不带小数
    startangle=90,
    textprops={'fontsize': 10, 'weight': 'bold', 'color': 'white'}  # 百分比文字样式
)

# 设置百分比文字背景为黑色矩形
for autotext in autotexts:
    autotext.set_bbox(dict(boxstyle='round,pad=0.3', facecolor='black', edgecolor='none', alpha=0.7))

# 添加图例在右侧
ax.legend(wedges, labels, loc="center left", bbox_to_anchor=(1, 0, 0.5, 1), fontsize=10)

# 设置标题
plt.title("磁盘告警", fontsize=14, fontweight='bold', pad=20)

# 确保饼图是圆的
ax.set_aspect('equal')

# 调整布局，为图例留出空间
plt.tight_layout()

# 显示图形
plt.show()
