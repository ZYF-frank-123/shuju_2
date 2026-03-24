import pandas as pd
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
matplotlib.rcParams['axes.unicode_minus'] = False

# 读取数据
df = pd.read_excel('教育机构学员缴费数据.xlsx')

# 将缴费日期转换为日期格式
df['缴费日期'] = pd.to_datetime(df['缴费日期'])

# ===================== 一、数据处理 =====================

print("=" * 60)
print("一、数据处理")
print("=" * 60)

# 1. 数据校验与修正
# 计算理论缴费总额
df['理论缴费总额'] = df['报名节数'] * df['单价（元/节）']

# 找出不一致的记录
inconsistent = df[df['缴费总额（元）'] != df['理论缴费总额']].copy()

if len(inconsistent) > 0:
    print("\n1. 缴费总额修正记录：")
    print("-" * 60)
    print(f"{'学员ID':<10} {'修正前':<10} {'修正后':<10} {'差异':<10}")
    print("-" * 60)
    for idx, row in inconsistent.iterrows():
        print(f"{row['学员ID']:<10} {row['缴费总额（元）']:<10} {row['理论缴费总额']:<10} {row['理论缴费总额'] - row['缴费总额（元）']:<10}")
else:
    print("\n1. 缴费总额校验：所有记录一致，无需修正")

# 修正缴费总额
df['缴费总额（元）'] = df['理论缴费总额']
df = df.drop('理论缴费总额', axis=1)

# 输出修正后的数据表（前10行示例）
print("\n2. 修正后的数据表（前10行）：")
print("-" * 120)
print(df.head(10).to_string())

# 2. 月度聚合汇总
# 提取月份
df['月份'] = df['缴费日期'].dt.strftime('%Y-%m')

# 「月度 - 课程类别 - 缴费总额」汇总表
monthly_course = df.groupby(['月份', '课程类别'])['缴费总额（元）'].sum().reset_index()
monthly_course = monthly_course.rename(columns={'缴费总额（元）': '缴费总额'})

print("\n3. 「月度 - 课程类别 - 缴费总额」汇总表：")
print("-" * 50)
print(monthly_course.to_string(index=False))

# 「月度 - 学员所在城市 - 缴费总额」汇总表
monthly_city = df.groupby(['月份', '学员所在城市'])['缴费总额（元）'].sum().reset_index()
monthly_city = monthly_city.rename(columns={'缴费总额（元）': '缴费总额', '学员所在城市': '城市'})

print("\n4. 「月度 - 学员所在城市 - 缴费总额」汇总表：")
print("-" * 50)
print(monthly_city.to_string(index=False))

# ===================== 二、统计分析 =====================

print("\n" + "=" * 60)
print("二、统计分析")
print("=" * 60)

total_payment = df['缴费总额（元）'].sum()
print(f"\n总缴费额：{total_payment:.2f} 元")

# 1. 维度占比计算
# 各城市缴费额占比
city_payment = df.groupby('学员所在城市')['缴费总额（元）'].sum().reset_index()
city_payment['占比(%)'] = (city_payment['缴费总额（元）'] / total_payment * 100).round(2)
city_payment = city_payment.sort_values('缴费总额（元）', ascending=False)

print("\n1. 各城市缴费额占比：")
print("-" * 50)
print(f"{'城市':<10} {'缴费总额(元)':<15} {'占比(%)':<10}")
print("-" * 50)
for idx, row in city_payment.iterrows():
    print(f"{row['学员所在城市']:<10} {row['缴费总额（元）']:<15.2f} {row['占比(%)']:<10.2f}")

# 各课程类别占比
category_payment = df.groupby('课程类别')['缴费总额（元）'].sum().reset_index()
category_payment['占比(%)'] = (category_payment['缴费总额（元）'] / total_payment * 100).round(2)
category_payment = category_payment.sort_values('缴费总额（元）', ascending=False)

print("\n2. 各课程类别缴费额占比：")
print("-" * 60)
print(f"{'课程类别':<15} {'缴费总额(元)':<15} {'占比(%)':<10}")
print("-" * 60)
for idx, row in category_payment.iterrows():
    print(f"{row['课程类别']:<15} {row['缴费总额（元）']:<15.2f} {row['占比(%)']:<10.2f}")

# 2. 课程TOP3评选
course_payment = df.groupby('课程名称')['缴费总额（元）'].sum().reset_index()
course_payment['贡献度(%)'] = (course_payment['缴费总额（元）'] / total_payment * 100).round(2)
course_payment = course_payment.sort_values('缴费总额（元）', ascending=False).head(3)

print("\n3. 缴费总额TOP3课程：")
print("-" * 60)
print(f"{'排名':<5} {'课程名称':<15} {'缴费总额(元)':<15} {'贡献度(%)':<10}")
print("-" * 60)
for i, (idx, row) in enumerate(course_payment.iterrows(), 1):
    print(f"{i:<5} {row['课程名称']:<15} {row['缴费总额（元）']:<15.2f} {row['贡献度(%)']:<10.2f}")

# 3. 多维度特征分析
# 课程状态分析
status_payment = df.groupby('课程状态')['缴费总额（元）'].sum().reset_index()
status_payment['占比(%)'] = (status_payment['缴费总额（元）'] / total_payment * 100).round(2)

print("\n4. 课程状态缴费额占比：")
print("-" * 60)
print(f"{'课程状态':<10} {'缴费总额(元)':<15} {'占比(%)':<10}")
print("-" * 60)
for idx, row in status_payment.iterrows():
    print(f"{row['课程状态']:<10} {row['缴费总额（元）']:<15.2f} {row['占比(%)']:<10.2f}")

# 单独统计已退费金额及退费占比
refund_payment = df[df['课程状态'] == '已退费']['缴费总额（元）'].sum()
refund_ratio = (refund_payment / total_payment * 100).round(2)
print(f"\n已退费总额：{refund_payment:.2f} 元，退费占比：{refund_ratio:.2f}%")

# 城市人均缴费分析
# 按学员ID去重统计各城市学员数量
city_students = df.drop_duplicates('学员ID').groupby('学员所在城市')['学员ID'].count().reset_index()
city_students = city_students.rename(columns={'学员ID': '学员数量'})

# 合并计算人均缴费
city_analysis = pd.merge(city_payment, city_students, on='学员所在城市')
city_analysis['人均缴费额(元)'] = (city_analysis['缴费总额（元）'] / city_analysis['学员数量']).round(2)
city_analysis = city_analysis.sort_values('人均缴费额(元)', ascending=False)

print("\n5. 城市人均缴费分析：")
print("-" * 80)
print(f"{'城市':<10} {'缴费总额(元)':<15} {'学员数量':<10} {'人均缴费额(元)':<15} {'占比(%)':<10}")
print("-" * 80)
for idx, row in city_analysis.iterrows():
    print(f"{row['学员所在城市']:<10} {row['缴费总额（元）']:<15.2f} {row['学员数量']:<10} {row['人均缴费额(元)']:<15.2f} {row['占比(%)']:<10.2f}")

# ===================== 三、可视化说明 =====================

print("\n" + "=" * 60)
print("三、可视化说明")
print("=" * 60)

print("\n1. 折线图：展示「月度缴费总额趋势」")
print("-" * 60)
print("适用场景：")
print("  ① 教培机构月度运营复盘会议，用于回顾整体营收表现")
print("  ② 招生旺季（如寒暑假、开学季）前后的营收波动分析")
print("  ③ 新课程推广效果的月度跟踪评估")
print("\n业务价值：")
print("  ① 直观识别营收高峰期和低谷期，帮助机构合理调配师资和资源")
print("  ② 预测未来营收趋势，为制定招生计划和预算提供数据支持")
print("  ③ 对比历年同期数据，评估机构业务增长或衰退情况")

print("\n2. 漏斗图：展示「学员所在城市→课程类别→课程状态」的缴费额分布")
print("-" * 60)
print("适用场景：")
print("  ① 机构区域市场拓展战略制定会议，分析各城市贡献度差异")
print("  ② 课程品类优化分析，识别高潜力课程类别")
print("  ③ 学员生命周期价值分析，从城市、课程到最终转化的全链路洞察")
print("\n业务价值：")
print("  ① 清晰展示缴费额在城市、课程类别、课程状态各层级的流失情况")
print("  ② 识别高价值城市和课程类别，为资源倾斜提供决策依据")
print("  ③ 发现课程状态转化瓶颈（如高退费课程），针对性优化教学服务")

# ===================== 四、可视化图表生成 =====================

print("\n" + "=" * 60)
print("四、可视化图表生成")
print("=" * 60)

# 1. 折线图：月度缴费总额趋势
print("\n正在生成折线图：月度缴费总额趋势...")
monthly_total = df.groupby('月份')['缴费总额（元）'].sum().reset_index()

plt.figure(figsize=(12, 6))
plt.plot(monthly_total['月份'], monthly_total['缴费总额（元）'], marker='o', linewidth=2, color='#2E86AB', markersize=8)
plt.title('月度缴费总额趋势图', fontsize=16, pad=20)
plt.xlabel('月份', fontsize=12)
plt.ylabel('缴费总额（元）', fontsize=12)
plt.grid(True, alpha=0.3)
plt.xticks(rotation=45)

# 添加数据标签
for x, y in zip(monthly_total['月份'], monthly_total['缴费总额（元）']):
    plt.text(x, y + 5000, f'{y:,.0f}', ha='center', fontsize=10)

plt.tight_layout()
plt.savefig('月度缴费总额趋势图.png', dpi=300, bbox_inches='tight')
plt.close()
print("折线图已保存：月度缴费总额趋势图.png")

# 2. 漏斗图：学员所在城市→课程类别→课程状态 缴费额分布
print("\n正在生成漏斗图：缴费额多维度分布...")

# 计算各层级数据
city_total = df.groupby('学员所在城市')['缴费总额（元）'].sum().sort_values(ascending=False).head(5)
category_total = df.groupby('课程类别')['缴费总额（元）'].sum().sort_values(ascending=False)
status_total = df.groupby('课程状态')['缴费总额（元）'].sum().sort_values(ascending=False)

# 创建漏斗图数据（使用百分比展示）
funnel_data = [
    ('总缴费额', total_payment),
    ('Top5城市合计', city_total.sum()),
    ('Top课程类别', category_total.iloc[0]),
    ('有效缴费（扣除退费）', total_payment - refund_payment)
]

labels = [x[0] for x in funnel_data]
values = [x[1] for x in funnel_data]
percentages = [v / total_payment * 100 for v in values]

plt.figure(figsize=(10, 8))
colors = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D']

# 绘制漏斗图
for i, (label, value, pct, color) in enumerate(zip(labels, values, percentages, colors)):
    width = pct / 100 * 0.8 + 0.2
    plt.barh(y=len(labels)-1-i, width=width, height=0.6, color=color, alpha=0.8)
    plt.text(width + 0.02, len(labels)-1-i, f'{label}\n{value:,.0f}元 ({pct:.1f}%)', va='center', fontsize=11)

plt.xlim(0, 1.1)
plt.ylim(-0.5, len(labels)-0.5)
plt.title('缴费额转化漏斗图', fontsize=16, pad=20)
plt.xticks([])
plt.yticks([])
plt.box(False)

plt.tight_layout()
plt.savefig('缴费额转化漏斗图.png', dpi=300, bbox_inches='tight')
plt.close()
print("漏斗图已保存：缴费额转化漏斗图.png")

# 3. 附加可视化：各城市缴费额柱状图
print("\n正在生成柱状图：各城市缴费额对比...")
plt.figure(figsize=(12, 6))
city_payment_sorted = city_payment.sort_values('缴费总额（元）', ascending=True)
bars = plt.barh(city_payment_sorted['学员所在城市'], city_payment_sorted['缴费总额（元）'], color='#2E86AB', alpha=0.7)

plt.title('各城市缴费总额对比', fontsize=16, pad=20)
plt.xlabel('缴费总额（元）', fontsize=12)

# 添加数据标签
for bar in bars:
    width = bar.get_width()
    plt.text(width + 5000, bar.get_y() + bar.get_height()/2, f'{width:,.0f}', ha='left', va='center')

plt.tight_layout()
plt.savefig('各城市缴费总额对比.png', dpi=300, bbox_inches='tight')
plt.close()
print("柱状图已保存：各城市缴费总额对比.png")

# 4. 附加可视化：课程类别占比饼图
print("\n正在生成饼图：课程类别缴费额占比...")
plt.figure(figsize=(8, 8))
colors = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D']
wedges, texts, autotexts = plt.pie(category_payment['缴费总额（元）'], 
                                   labels=category_payment['课程类别'],
                                   colors=colors,
                                   autopct='%1.1f%%',
                                   startangle=90)
plt.title('课程类别缴费额占比', fontsize=16, pad=20)
plt.setp(texts, fontsize=12)
plt.setp(autotexts, fontsize=12, weight='bold')

plt.tight_layout()
plt.savefig('课程类别缴费额占比.png', dpi=300, bbox_inches='tight')
plt.close()
print("饼图已保存：课程类别缴费额占比.png")

# 保存结果到Excel文件
with pd.ExcelWriter('分析结果汇总.xlsx') as writer:
    # 修正后的数据表
    df.to_excel(writer, sheet_name='修正后原始数据', index=False)
    # 月度-课程类别汇总
    monthly_course.to_excel(writer, sheet_name='月度课程类别汇总', index=False)
    # 月度-城市汇总
    monthly_city.to_excel(writer, sheet_name='月度城市汇总', index=False)
    # 城市占比
    city_payment.to_excel(writer, sheet_name='城市缴费占比', index=False)
    # 课程类别占比
    category_payment.to_excel(writer, sheet_name='课程类别占比', index=False)
    # TOP3课程
    course_payment.to_excel(writer, sheet_name='TOP3课程', index=False)
    # 课程状态
    status_payment.to_excel(writer, sheet_name='课程状态分析', index=False)
    # 城市人均缴费
    city_analysis.to_excel(writer, sheet_name='城市人均缴费', index=False)

print("\n" + "=" * 60)
print("分析完成！结果已保存至：分析结果汇总.xlsx")
print("=" * 60)
