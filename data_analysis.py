import pandas as pd
import numpy as np
from datetime import datetime

# 读取Excel数据
df = pd.read_excel(r'f:\github\shujufenxi_2\g\教育机构学员缴费数据.xlsx')

print("="*80)
print("原始数据前10行：")
print(df.head(10))
print("\n数据列名：", df.columns.tolist())
print("\n数据形状：", df.shape)

# 检查列名
print("\n各列数据类型：")
print(df.dtypes)

# 计算理论缴费总额
df['理论缴费总额'] = df['报名节数'] * df['单价（元/节）']

# 校验不一致的记录
mismatched = df[df['缴费总额（元）'] != df['理论缴费总额']]

print("\n" + "="*80)
print("一、数据校验与修正")
print("="*80)

if len(mismatched) > 0:
    print(f"\n发现 {len(mismatched)} 条计算不一致的记录：")
    print("\n修正记录明细：")
    corrections = []
    for idx, row in mismatched.iterrows():
        student_id = row['学员ID']
        old_value = row['缴费总额（元）']
        new_value = row['理论缴费总额']
        corrections.append({
            '学员ID': student_id,
            '修正前缴费总额': old_value,
            '修正后缴费总额': new_value
        })
        print(f"  学员ID: {student_id}, 修正前: {old_value}, 修正后: {new_value}")
    
    # 修正数据
    df.loc[df['缴费总额（元）'] != df['理论缴费总额'], '缴费总额（元）'] = df.loc[df['缴费总额（元）'] != df['理论缴费总额'], '理论缴费总额']
    
    # 输出修正后的数据表
    print("\n修正后的数据表（可复制到Excel）：")
    corrected_df = df[['学员ID', '学员姓名', '课程类别', '课程名称', '报名节数', '单价（元/节）', '缴费总额（元）', '缴费日期', '学员所在城市', '缴费方式', '课程状态']].copy()
    print(corrected_df.to_string(index=False))
else:
    print("\n所有记录计算一致，无需修正。")

# 保存修正后的数据
corrected_df = df[['学员ID', '学员姓名', '课程类别', '课程名称', '报名节数', '单价（元/节）', '缴费总额（元）', '缴费日期', '学员所在城市', '缴费方式', '课程状态']].copy()
corrected_df.to_excel('修正后数据表.xlsx', index=False)
print("\n修正后数据表已保存至：修正后数据表.xlsx")

# 提取月份
df['月份'] = pd.to_datetime(df['缴费日期']).dt.to_period('M')

print("\n" + "="*80)
print("二、月度聚合汇总")
print("="*80)

# 月度-课程类别-缴费总额汇总表
summary_category = df.groupby(['月份', '课程类别'])['缴费总额（元）'].sum().reset_index()
summary_category.columns = ['月份', '课程类别', '缴费总额']
print("\n【月度-课程类别-缴费总额汇总表】")
print(summary_category.to_string(index=False))
summary_category.to_excel('月度-课程类别-缴费总额汇总表.xlsx', index=False)

# 月度-学员所在城市-缴费总额汇总表
summary_city = df.groupby(['月份', '学员所在城市'])['缴费总额（元）'].sum().reset_index()
summary_city.columns = ['月份', '城市', '缴费总额']
print("\n【月度-学员所在城市-缴费总额汇总表】")
print(summary_city.to_string(index=False))
summary_city.to_excel('月度-学员所在城市-缴费总额汇总表.xlsx', index=False)

print("\n汇总表已保存。")

print("\n" + "="*80)
print("三、统计分析")
print("="*80)

# 总缴费额
total_amount = df['缴费总额（元）'].sum()
print(f"\n总缴费额：{total_amount} 元")

# 1. 各城市缴费额占比
city_amount = df.groupby('学员所在城市')['缴费总额（元）'].sum().sort_values(ascending=False)
city_ratio = (city_amount / total_amount * 100).round(2)
print("\n【各城市缴费额占比】")
city_stats = pd.DataFrame({
    '城市': city_amount.index,
    '缴费总额（元）': city_amount.values,
    '占比（%）': city_ratio.values
})
print(city_stats.to_string(index=False))

# 2. 各课程类别缴费额占比
category_amount = df.groupby('课程类别')['缴费总额（元）'].sum().sort_values(ascending=False)
category_ratio = (category_amount / total_amount * 100).round(2)
print("\n【各课程类别缴费额占比】")
category_stats = pd.DataFrame({
    '课程类别': category_amount.index,
    '缴费总额（元）': category_amount.values,
    '占比（%）': category_ratio.values
})
print(category_stats.to_string(index=False))

# 3. 课程TOP3
course_amount = df.groupby('课程名称')['缴费总额（元）'].sum().sort_values(ascending=False)
top3_courses = course_amount.head(3)
top3_ratio = (top3_courses / total_amount * 100).round(2)
print("\n【课程TOP3评选】")
print("排名 | 课程名称 | 缴费总额（元） | 贡献度（%）")
for i, (course, amount) in enumerate(top3_courses.items(), 1):
    print(f"  {i}  | {course} | {amount} | {top3_ratio[course]}")

# 4. 课程状态分析
status_amount = df.groupby('课程状态')['缴费总额（元）'].sum().sort_values(ascending=False)
status_ratio = (status_amount / total_amount * 100).round(2)
refund_amount = status_amount.get('已退费', 0)
refund_ratio = (refund_amount / total_amount * 100).round(2)
print("\n【课程状态分析】")
status_stats = pd.DataFrame({
    '课程状态': status_amount.index,
    '缴费总额（元）': status_amount.values,
    '占比（%）': status_ratio.values
})
print(status_stats.to_string(index=False))
print(f"\n已退费金额：{refund_amount} 元，退费占比：{refund_ratio}%")

# 5. 城市人均缴费分析
city_students = df.groupby('学员所在城市')['学员ID'].nunique()
city_avg = (city_amount / city_students).round(2)
print("\n【城市人均缴费分析】")
print("城市 | 学员数量 | 城市总缴费额（元） | 人均缴费额（元）")
for city in city_amount.index:
    print(f"{city} | {city_students[city]} | {city_amount[city]} | {city_avg[city]}")

# 保存所有统计结果到Excel
with pd.ExcelWriter('统计分析结果.xlsx') as writer:
    city_stats.to_excel(writer, sheet_name='城市缴费占比', index=False)
    category_stats.to_excel(writer, sheet_name='课程类别占比', index=False)
    status_stats.to_excel(writer, sheet_name='课程状态分析', index=False)
    
    # TOP3课程
    top3_df = pd.DataFrame({
        '排名': [1, 2, 3],
        '课程名称': top3_courses.index,
        '缴费总额（元）': top3_courses.values,
        '贡献度（%）': top3_ratio.values
    })
    top3_df.to_excel(writer, sheet_name='课程TOP3', index=False)
    
    # 城市人均
    city_avg_df = pd.DataFrame({
        '城市': city_amount.index,
        '学员数量': city_students.values,
        '城市总缴费额（元）': city_amount.values,
        '人均缴费额（元）': city_avg.values
    })
    city_avg_df.to_excel(writer, sheet_name='城市人均缴费', index=False)

print("\n统计分析结果已保存至：统计分析结果.xlsx")
