import pandas as pd
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
matplotlib.rcParams['axes.unicode_minus'] = False

# 读取Excel数据
df = pd.read_excel('教育机构学员缴费数据.xlsx')

# 查看数据基本信息
print("=" * 80)
print("数据基本信息")
print("=" * 80)
print(f"数据行数: {len(df)}")
print(f"数据列: {df.columns.tolist()}")
print("\n前5行数据:")
print(df.head())
print("\n数据类型:")
print(df.dtypes)

# 数据校验与修正
print("\n" + "=" * 80)
print("一、数据校验与修正")
print("=" * 80)

# 计算正确的缴费总额
df['计算缴费总额'] = df['报名节数'] * df['单价（元/节）']

# 找出不一致的记录
df['金额一致'] = abs(df['缴费总额（元）'] - df['计算缴费总额']) < 0.01  # 允许0.01元的浮点误差
inconsistent = df[~df['金额一致']].copy()

print(f"\n发现 {len(inconsistent)} 条记录缴费总额计算不一致")

if len(inconsistent) > 0:
    print("\n修正记录明细:")
    print("-" * 100)
    print(f"{'学员ID':<12} {'学员姓名':<10} {'原缴费总额':<12} {'修正后缴费总额':<14} {'差异':<12}")
    print("-" * 100)
    
    for idx, row in inconsistent.iterrows():
        original = row['缴费总额（元）']
        corrected = row['计算缴费总额']
        diff = corrected - original
        print(f"{row['学员ID']:<12} {row['学员姓名']:<10} {original:<12.2f} {corrected:<14.2f} {diff:<12.2f}")
    
    # 修正缴费总额
    df['缴费总额（元）_修正'] = df['计算缴费总额']
else:
    df['缴费总额（元）_修正'] = df['缴费总额（元）']

print(f"\n所有记录已修正完成，修正后数据总行数: {len(df)}")

# 提取月份
df['缴费日期'] = pd.to_datetime(df['缴费日期'])
df['月份'] = df['缴费日期'].dt.to_period('M').astype(str)

# 保存修正后的数据
df_corrected = df[['学员ID', '学员姓名', '课程类别', '课程名称', '报名节数', 
                   '单价（元/节）', '缴费总额（元）', '缴费总额（元）_修正', 
                   '缴费日期', '学员所在城市', '缴费方式', '课程状态', '月份']].copy()

df_corrected.to_excel('修正后的数据表.xlsx', index=False)
print("\n修正后的数据已保存至: 修正后的数据表.xlsx")

# 月度聚合汇总
print("\n" + "=" * 80)
print("二、月度聚合汇总表")
print("=" * 80)

# 汇总表1：月度-课程类别-缴费总额
monthly_course = df.groupby(['月份', '课程类别'])['缴费总额（元）_修正'].sum().reset_index()
monthly_course.columns = ['月份', '课程类别', '缴费总额']
monthly_course['缴费总额'] = monthly_course['缴费总额'].round(2)
monthly_course.to_excel('月度-课程类别-缴费总额汇总表.xlsx', index=False)

print("\n【汇总表1】月度-课程类别-缴费总额:")
print("-" * 60)
print(monthly_course.to_string(index=False))

# 汇总表2：月度-学员所在城市-缴费总额
monthly_city = df.groupby(['月份', '学员所在城市'])['缴费总额（元）_修正'].sum().reset_index()
monthly_city.columns = ['月份', '城市', '缴费总额']
monthly_city['缴费总额'] = monthly_city['缴费总额'].round(2)
monthly_city.to_excel('月度-学员所在城市-缴费总额汇总表.xlsx', index=False)

print("\n【汇总表2】月度-学员所在城市-缴费总额:")
print("-" * 60)
print(monthly_city.to_string(index=False))

# 统计分析
print("\n" + "=" * 80)
print("三、统计分析")
print("=" * 80)

total_amount = df['缴费总额（元）_修正'].sum()
print(f"\n总缴费额: {total_amount:,.2f} 元")

# 1. 维度占比计算
print("\n" + "-" * 60)
print("1. 各维度占比分析")
print("-" * 60)

# 城市占比
city_stats = df.groupby('学员所在城市')['缴费总额（元）_修正'].sum().reset_index()
city_stats.columns = ['城市', '缴费总额']
city_stats['占比(%)'] = (city_stats['缴费总额'] / total_amount * 100).round(2)
city_stats = city_stats.sort_values('缴费总额', ascending=False)
city_stats['排名'] = range(1, len(city_stats) + 1)

print("\n【城市缴费额占比】")
print(city_stats.to_string(index=False))

# 课程类别占比
category_stats = df.groupby('课程类别')['缴费总额（元）_修正'].sum().reset_index()
category_stats.columns = ['课程类别', '缴费总额']
category_stats['占比(%)'] = (category_stats['缴费总额'] / total_amount * 100).round(2)
category_stats = category_stats.sort_values('缴费总额', ascending=False)
category_stats['排名'] = range(1, len(category_stats) + 1)

print("\n【课程类别缴费额占比】")
print(category_stats.to_string(index=False))

# 2. 课程TOP3评选
print("\n" + "-" * 60)
print("2. 课程TOP3评选")
print("-" * 60)

course_stats = df.groupby('课程名称')['缴费总额（元）_修正'].sum().reset_index()
course_stats.columns = ['课程名称', '缴费总额']
course_stats['贡献度(%)'] = (course_stats['缴费总额'] / total_amount * 100).round(2)
course_stats = course_stats.sort_values('缴费总额', ascending=False)
course_stats['排名'] = range(1, len(course_stats) + 1)

top3 = course_stats.head(3)
print("\n【TOP3课程】")
print(top3.to_string(index=False))

# 3. 课程状态分析
print("\n" + "-" * 60)
print("3. 课程状态分析")
print("-" * 60)

status_stats = df.groupby('课程状态')['缴费总额（元）_修正'].sum().reset_index()
status_stats.columns = ['课程状态', '缴费总额']
status_stats['占比(%)'] = (status_stats['缴费总额'] / total_amount * 100).round(2)
status_stats = status_stats.sort_values('缴费总额', ascending=False)

print("\n【各课程状态缴费额占比】")
print(status_stats.to_string(index=False))

# 退费统计
refund_amount = df[df['课程状态'] == '已退费']['缴费总额（元）_修正'].sum()
refund_ratio = (refund_amount / total_amount * 100)

print(f"\n【退费专项统计】")
print(f"已退费总额: {refund_amount:,.2f} 元")
print(f"退费占比: {refund_ratio:.2f}%")

# 4. 城市人均缴费分析
print("\n" + "-" * 60)
print("4. 城市人均缴费分析")
print("-" * 60)

# 按学员ID去重统计各城市学员数量
city_student_count = df.drop_duplicates(subset=['学员ID', '学员所在城市']).groupby('学员所在城市').size().reset_index(name='学员数量')
city_total = df.groupby('学员所在城市')['缴费总额（元）_修正'].sum().reset_index()
city_total.columns = ['学员所在城市', '总缴费额']

city_per_capita = pd.merge(city_total, city_student_count, on='学员所在城市')
city_per_capita['人均缴费额'] = (city_per_capita['总缴费额'] / city_per_capita['学员数量']).round(2)
city_per_capita = city_per_capita.sort_values('人均缴费额', ascending=False)
city_per_capita['排名'] = range(1, len(city_per_capita) + 1)

print("\n【城市人均缴费分析】")
print(city_per_capita.to_string(index=False))

# 保存统计分析结果到Excel
with pd.ExcelWriter('统计分析结果.xlsx') as writer:
    city_stats.to_excel(writer, sheet_name='城市占比', index=False)
    category_stats.to_excel(writer, sheet_name='课程类别占比', index=False)
    top3.to_excel(writer, sheet_name='TOP3课程', index=False)
    status_stats.to_excel(writer, sheet_name='课程状态分析', index=False)
    city_per_capita.to_excel(writer, sheet_name='城市人均缴费', index=False)

print("\n" + "=" * 80)
print("统计分析结果已保存至: 统计分析结果.xlsx")
print("=" * 80)

# 四、可视化图表生成
print("\n" + "=" * 80)
print("四、可视化图表生成")
print("=" * 80)

# 1. 折线图：月度缴费总额趋势
monthly_total = df.groupby('月份')['缴费总额（元）_修正'].sum().reset_index()
monthly_total.columns = ['月份', '缴费总额']

fig, ax = plt.subplots(figsize=(12, 6))
ax.plot(monthly_total['月份'], monthly_total['缴费总额'], marker='o', linewidth=2, markersize=8, color='#2E86AB')
ax.fill_between(monthly_total['月份'], monthly_total['缴费总额'], alpha=0.3, color='#2E86AB')
ax.set_xlabel('月份', fontsize=12)
ax.set_ylabel('缴费总额（元）', fontsize=12)
ax.set_title('月度缴费总额趋势图', fontsize=14, fontweight='bold')
ax.grid(True, alpha=0.3)
ax.tick_params(axis='x', rotation=45)

# 添加数值标签
for i, row in monthly_total.iterrows():
    ax.annotate(f'{row["缴费总额"]:,.0f}', 
                xy=(row['月份'], row['缴费总额']), 
                xytext=(0, 10), 
                textcoords='offset points',
                ha='center', fontsize=9)

plt.tight_layout()
plt.savefig('月度缴费总额趋势图.png', dpi=300, bbox_inches='tight')
print("\n✅ 月度缴费总额趋势图已保存: 月度缴费总额趋势图.png")

# 2. 多维度漏斗图展示
fig, axes = plt.subplots(1, 3, figsize=(18, 6))

# 漏斗图1：城市维度漏斗
city_funnel = city_stats.sort_values('缴费总额', ascending=True).tail(8)
colors1 = plt.cm.Blues(np.linspace(0.4, 0.9, len(city_funnel)))
bars1 = axes[0].barh(city_funnel['城市'], city_funnel['缴费总额'], color=colors1)
axes[0].set_xlabel('缴费总额（元）', fontsize=11)
axes[0].set_title('城市缴费额分布（漏斗图）', fontsize=12, fontweight='bold')
axes[0].grid(axis='x', alpha=0.3)
for i, (bar, value) in enumerate(zip(bars1, city_funnel['缴费总额'])):
    axes[0].text(value + 5000, bar.get_y() + bar.get_height()/2, 
                f'{value:,.0f}\n({city_funnel.iloc[i]["占比(%)"]}%)', 
                va='center', fontsize=9)

# 漏斗图2：课程类别维度漏斗
category_funnel = category_stats.sort_values('缴费总额', ascending=True)
colors2 = plt.cm.Greens(np.linspace(0.4, 0.9, len(category_funnel)))
bars2 = axes[1].barh(category_funnel['课程类别'], category_funnel['缴费总额'], color=colors2)
axes[1].set_xlabel('缴费总额（元）', fontsize=11)
axes[1].set_title('课程类别缴费额分布（漏斗图）', fontsize=12, fontweight='bold')
axes[1].grid(axis='x', alpha=0.3)
for i, (bar, value) in enumerate(zip(bars2, category_funnel['缴费总额'])):
    axes[1].text(value + 5000, bar.get_y() + bar.get_height()/2, 
                f'{value:,.0f}\n({category_funnel.iloc[i]["占比(%)"]}%)', 
                va='center', fontsize=9)

# 漏斗图3：课程状态维度漏斗
status_funnel = status_stats.sort_values('缴费总额', ascending=True)
colors3 = ['#FF6B6B' if x == '已退费' else '#4ECDC4' if x == '已结课' else '#45B7D1' if x == '已开课' else '#96CEB4' for x in status_funnel['课程状态']]
bars3 = axes[2].barh(status_funnel['课程状态'], status_funnel['缴费总额'], color=colors3)
axes[2].set_xlabel('缴费总额（元）', fontsize=11)
axes[2].set_title('课程状态缴费额分布（漏斗图）', fontsize=12, fontweight='bold')
axes[2].grid(axis='x', alpha=0.3)
for i, (bar, value) in enumerate(zip(bars3, status_funnel['缴费总额'])):
    axes[2].text(value + 5000, bar.get_y() + bar.get_height()/2, 
                f'{value:,.0f}\n({status_funnel.iloc[i]["占比(%)"]}%)', 
                va='center', fontsize=9)

plt.tight_layout()
plt.savefig('多维度漏斗图.png', dpi=300, bbox_inches='tight')
print("✅ 多维度漏斗图已保存: 多维度漏斗图.png")

# 3. 综合仪表盘：月度-课程类别堆叠面积图
monthly_course_pivot = monthly_course.pivot(index='月份', columns='课程类别', values='缴费总额').fillna(0)
fig, ax = plt.subplots(figsize=(14, 7))
monthly_course_pivot.plot(kind='area', stacked=True, ax=ax, alpha=0.7, 
                          color=['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4'])
ax.set_xlabel('月份', fontsize=12)
ax.set_ylabel('缴费总额（元）', fontsize=12)
ax.set_title('月度-课程类别缴费额趋势（堆叠面积图）', fontsize=14, fontweight='bold')
ax.legend(title='课程类别', loc='upper left', bbox_to_anchor=(1, 1))
ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig('月度课程类别趋势图.png', dpi=300, bbox_inches='tight')
print("✅ 月度课程类别趋势图已保存: 月度课程类别趋势图.png")

# 4. TOP3课程贡献度饼图
fig, axes = plt.subplots(1, 2, figsize=(14, 6))

# TOP3课程饼图
top3_data = top3.copy()
other_amount = course_stats.iloc[3:]['缴费总额'].sum()
pie_data = list(top3_data['缴费总额']) + [other_amount]
pie_labels = list(top3_data['课程名称']) + ['其他课程']
colors_pie = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#CCCCCC']
explode = (0.05, 0.05, 0.05, 0)

axes[0].pie(pie_data, labels=pie_labels, autopct='%1.2f%%', startangle=90, 
            colors=colors_pie, explode=explode, shadow=True)
axes[0].set_title('TOP3课程贡献度分布', fontsize=12, fontweight='bold')

# 城市人均缴费柱状图
city_per_capita_sorted = city_per_capita.sort_values('人均缴费额', ascending=True)
colors_bar = plt.cm.RdYlGn(np.linspace(0.3, 0.9, len(city_per_capita_sorted)))
bars = axes[1].barh(city_per_capita_sorted['学员所在城市'], city_per_capita_sorted['人均缴费额'], color=colors_bar)
axes[1].set_xlabel('人均缴费额（元）', fontsize=11)
axes[1].set_title('各城市人均缴费额对比', fontsize=12, fontweight='bold')
axes[1].grid(axis='x', alpha=0.3)
for bar, value in zip(bars, city_per_capita_sorted['人均缴费额']):
    axes[1].text(value + 100, bar.get_y() + bar.get_height()/2, f'{value:,.2f}', va='center', fontsize=9)

plt.tight_layout()
plt.savefig('TOP3课程与城市人均缴费分析图.png', dpi=300, bbox_inches='tight')
print("✅ TOP3课程与城市人均缴费分析图已保存: TOP3课程与城市人均缴费分析图.png")

print("\n" + "=" * 80)
print("所有可视化图表生成完成！")
print("=" * 80)
print("\n生成的图表文件：")
print("  1. 月度缴费总额趋势图.png - 展示月度营收趋势")
print("  2. 多维度漏斗图.png - 城市/课程类别/课程状态三维度漏斗")
print("  3. 月度课程类别趋势图.png - 各课程类别月度趋势堆叠图")
print("  4. TOP3课程与城市人均缴费分析图.png - 课程贡献度和城市人均缴费对比")
print("=" * 80)
