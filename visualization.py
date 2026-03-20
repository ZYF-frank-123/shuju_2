import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np

plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

df = pd.read_excel(r'f:\github\shujufenxi_2\g\教育机构学员缴费数据.xlsx')
df['月份'] = pd.to_datetime(df['缴费日期']).dt.to_period('M')

fig, axes = plt.subplots(2, 2, figsize=(16, 12))

# 图1：月度缴费总额趋势折线图
monthly_total = df.groupby('月份')['缴费总额（元）'].sum()
ax1 = axes[0, 0]
months = [str(m) for m in monthly_total.index]
values = monthly_total.values
ax1.plot(months, values, marker='o', linewidth=2.5, markersize=10, color='#2E86AB', markerfacecolor='#E94F37')
ax1.fill_between(months, values, alpha=0.3, color='#2E86AB')
for i, (m, v) in enumerate(zip(months, values)):
    ax1.annotate(f'{v:,.0f}', (m, v), textcoords="offset points", xytext=(0, 12), ha='center', fontsize=10, fontweight='bold')
ax1.set_title('月度缴费总额趋势', fontsize=16, fontweight='bold', pad=15)
ax1.set_xlabel('月份', fontsize=12)
ax1.set_ylabel('缴费总额（元）', fontsize=12)
ax1.grid(True, linestyle='--', alpha=0.7)
ax1.tick_params(axis='x', rotation=45)

# 图2：漏斗图 - 城市→课程类别→课程状态
ax2 = axes[0, 1]
city_total = df.groupby('学员所在城市')['缴费总额（元）'].sum().sort_values(ascending=False)
category_total = df.groupby('课程类别')['缴费总额（元）'].sum().sort_values(ascending=False)
status_total = df.groupby('课程状态')['缴费总额（元）'].sum().sort_values(ascending=False)

funnel_data = []
for city, amount in city_total.items():
    funnel_data.append(('城市: ' + city, amount, '城市'))
for cat, amount in category_total.items():
    funnel_data.append(('课程类别: ' + cat, amount, '课程类别'))
for status, amount in status_total.items():
    funnel_data.append(('课程状态: ' + status, amount, '课程状态'))

max_amount = max([d[1] for d in funnel_data])
colors = {'城市': '#3498db', '课程类别': '#2ecc71', '课程状态': '#e74c3c'}
y_positions = list(range(len(funnel_data)))

for i, (label, amount, level) in enumerate(funnel_data):
    width = amount / max_amount * 0.8
    left = (1 - width) / 2
    color = colors[level]
    ax2.barh(y_positions[i], width, left=left, height=0.7, color=color, alpha=0.8, edgecolor='white', linewidth=1)
    ax2.text(0.5, y_positions[i], f'{label}\n{amount:,.0f}元', ha='center', va='center', fontsize=9, fontweight='bold', color='white')

ax2.set_xlim(0, 1)
ax2.set_ylim(-0.5, len(funnel_data) - 0.5)
ax2.invert_yaxis()
ax2.set_title('漏斗图：城市→课程类别→课程状态 缴费额分布', fontsize=14, fontweight='bold', pad=15)
ax2.axis('off')

legend_elements = [plt.Rectangle((0, 0), 1, 1, facecolor=colors['城市'], alpha=0.8, label='城市'),
                   plt.Rectangle((0, 0), 1, 1, facecolor=colors['课程类别'], alpha=0.8, label='课程类别'),
                   plt.Rectangle((0, 0), 1, 1, facecolor=colors['课程状态'], alpha=0.8, label='课程状态')]
ax2.legend(handles=legend_elements, loc='lower right', fontsize=10)

# 图3：各城市缴费额占比饼图
ax3 = axes[1, 0]
city_amount = df.groupby('学员所在城市')['缴费总额（元）'].sum().sort_values(ascending=False)
colors_pie = plt.cm.Set3(np.linspace(0, 1, len(city_amount)))
wedges, texts, autotexts = ax3.pie(city_amount.values, labels=city_amount.index, autopct='%1.1f%%',
                                    colors=colors_pie, startangle=90, explode=[0.02]*len(city_amount))
ax3.set_title('各城市缴费额占比', fontsize=16, fontweight='bold', pad=15)
for autotext in autotexts:
    autotext.set_fontsize(10)
    autotext.set_fontweight('bold')

# 图4：课程类别缴费额对比柱状图
ax4 = axes[1, 1]
category_amount = df.groupby('课程类别')['缴费总额（元）'].sum().sort_values(ascending=True)
bars = ax4.barh(category_amount.index, category_amount.values, color=['#3498db', '#2ecc71', '#f39c12', '#e74c3c'])
for bar, value in zip(bars, category_amount.values):
    ax4.text(value + 5000, bar.get_y() + bar.get_height()/2, f'{value:,.0f}元', va='center', fontsize=11, fontweight='bold')
ax4.set_title('各课程类别缴费总额', fontsize=16, fontweight='bold', pad=15)
ax4.set_xlabel('缴费总额（元）', fontsize=12)
ax4.grid(True, axis='x', linestyle='--', alpha=0.7)

plt.tight_layout(pad=3.0)
plt.savefig('可视化分析图表.png', dpi=150, bbox_inches='tight', facecolor='white')
plt.close()

print("可视化图表已保存至：可视化分析图表.png")

# 生成单独的折线图
fig2, ax = plt.subplots(figsize=(12, 6))
monthly_total = df.groupby('月份')['缴费总额（元）'].sum()
months = [str(m) for m in monthly_total.index]
values = monthly_total.values

ax.plot(months, values, marker='o', linewidth=3, markersize=12, color='#2E86AB', markerfacecolor='#E94F37', markeredgewidth=2, markeredgecolor='white')
ax.fill_between(months, values, alpha=0.2, color='#2E86AB')

for i, (m, v) in enumerate(zip(months, values)):
    ax.annotate(f'{v:,.0f}元', (m, v), textcoords="offset points", xytext=(0, 15), ha='center', fontsize=11, fontweight='bold', color='#E94F37')

ax.set_title('教育机构月度缴费总额趋势\n（2025年9月 - 2026年2月）', fontsize=18, fontweight='bold', pad=20)
ax.set_xlabel('月份', fontsize=14)
ax.set_ylabel('缴费总额（元）', fontsize=14)
ax.grid(True, linestyle='--', alpha=0.5)
ax.tick_params(axis='x', rotation=45)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

plt.tight_layout()
plt.savefig('月度缴费趋势折线图.png', dpi=150, bbox_inches='tight', facecolor='white')
plt.close()

print("月度缴费趋势折线图已保存至：月度缴费趋势折线图.png")
