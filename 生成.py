# edu_student_payment.py
import pandas as pd
import random
from datetime import datetime, timedelta

# 固定随机种子，确保每次生成数据完全一致
random.seed(42)

# 1. 基础配置
# 时间范围：2025-09-01 至 2026-02-28（规避2026年2月无29日的问题）
start_date = datetime(2025, 9, 1)
end_date = datetime(2026, 2, 28)
date_range = [start_date + timedelta(days=x) for x in range((end_date - start_date).days + 1)]
# 生成数据条数
num_rows = 200

# 2. 基础数据字典（贴合教培机构真实业务）
last_names = ["王", "李", "张", "刘", "陈", "杨", "赵", "黄", "周", "吴"]
first_names = ["小明", "小红", "小刚", "小丽", "小杰", "小芳", "小亮", "小敏"]
# 课程类目+对应课程名称（关联匹配，保证逻辑）
course_categories = ["K12文化课", "艺术特长", "职业技能", "语言培训"]
course_names = {
    "K12文化课": ["小学数学", "初中英语", "高中物理", "小学语文"],
    "艺术特长": ["钢琴", "舞蹈", "绘画", "书法"],
    "职业技能": ["Python编程", "UI设计", "会计考证", "电商运营"],
    "语言培训": ["成人英语", "日语考级", "雅思", "托福"]
}
# 学员所在城市
cities = ["北京", "上海", "广州", "深圳", "杭州", "南京", "成都", "武汉"]
# 缴费方式
pay_methods = ["微信支付", "支付宝", "银行卡", "现金", "分期"]
# 课程状态
course_status = ["已报名", "已开课", "已结课", "已退费"]

# 3. 生成数据
data = []
for i in range(num_rows):
    # 学员ID（格式：XY001-XY200）
    stu_id = f"XY{i + 1:03d}"
    # 学员姓名（随机组合）
    stu_name = random.choice(last_names) + random.choice(first_names)
    # 课程类目+课程名称（关联匹配）
    category = random.choice(course_categories)
    course = random.choice(course_names[category])
    # 报名节数（10-60节，贴合教培机构常规报名量）
    sign_times = random.randint(10, 60)
    # 单价（按课程类目分级，符合市场行情）
    if category == "K12文化课":
        price_per_class = random.randint(100, 200)
    elif category == "艺术特长":
        price_per_class = random.randint(150, 300)
    elif category == "职业技能":
        price_per_class = random.randint(200, 400)
    else:  # 语言培训
        price_per_class = random.randint(180, 350)
    # 缴费总额（自动计算：报名节数 × 单价）
    total_fee = sign_times * price_per_class
    # 缴费日期（随机选近6个月某天）
    pay_date = random.choice(date_range).strftime("%Y-%m-%d")
    # 学员所在城市/缴费方式/课程状态
    city = random.choice(cities)
    pay_method = random.choice(pay_methods)
    status = random.choice(course_status)

    # 组装单条数据
    data.append([
        stu_id, stu_name, category, course, sign_times, price_per_class,
        total_fee, pay_date, city, pay_method, status
    ])

# 4. 写入Excel文件
columns = [
    "学员ID", "学员姓名", "课程类别", "课程名称", "报名节数", "单价（元/节）",
    "缴费总额（元）", "缴费日期", "学员所在城市", "缴费方式", "课程状态"
]
df = pd.DataFrame(data, columns=columns)
# 写入Excel，index=False不生成行号，engine="openpyxl"支持.xlsx格式
df.to_excel("教育机构学员缴费数据.xlsx", index=False, engine="openpyxl")

# 输出生成结果提示
print(f"教育机构学员缴费数据生成完成！")
print(f"数据条数：{len(df)} 条")
print(f"文件保存路径：当前目录下「教育机构学员缴费数据.xlsx」")