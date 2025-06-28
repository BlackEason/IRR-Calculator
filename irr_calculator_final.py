import pandas as pd
import numpy as np

def calculate_irr(cash_flows):
    """IRR计算函数"""
    def npv(rate):
        return sum([cf / (1 + rate) ** i for i, cf in enumerate(cash_flows)])
    
    low, high = -0.99, 5.0
    for _ in range(1000):
        mid = (low + high) / 2
        if abs(npv(mid)) < 0.0001:
            return mid
        elif npv(mid) > 0:
            low = mid
        else:
            high = mid
    return mid

print("=== IRR计算器 - 最终版 ===")
print("显示所有数据行，IRR为负数时显示为空")
print()

# 读取数据
try:
    df_full = pd.read_excel("数据底稿.xlsx", sheet_name="Sheet1", header=None)
    print("✅ 成功读取数据底稿")
except Exception as e:
    print(f"❌ 读取文件失败: {e}")
    exit()

# 查找数据开始行
data_start_row = None
for i in range(len(df_full)):
    if pd.notna(df_full.iloc[i, 0]) and str(df_full.iloc[i, 0]).strip() == '1':
        data_start_row = i
        break

if data_start_row is None:
    print("❌ 未找到数据开始行")
    exit()

print(f"数据从第{data_start_row + 1}行开始")

# 读取实际数据
df = df_full.iloc[data_start_row:].reset_index(drop=True)
df.columns = ['保单年度', '提取金额', '提取后保证金额', '提取后复归红利', '提取后终期红利', '提取后总价值', '预期年化复利']

# 从数据底稿提取保费信息
print("\n=== 保费缴费信息 ===")
premium_info = {
    '缴费年限': df_full.iloc[1, 1],
    '年缴金额': df_full.iloc[2, 1],
    '首年保费优惠': df_full.iloc[3, 1],
    '优惠后': df_full.iloc[4, 1],
    '保险业监管局征费': df_full.iloc[5, 1],
    '首年保费终': df_full.iloc[6, 1],
}

for key, value in premium_info.items():
    print(f"  {key}: {value}")

# 构建保费现金流
pay_years = int(premium_info['缴费年限'])
annual_premium = float(premium_info['年缴金额'])
first_year_final = float(premium_info['首年保费终'])

premiums = [-first_year_final]
for i in range(1, pay_years):
    premiums.append(-annual_premium)

total_premium = -sum(premiums)
print(f"\n保费现金流: {premiums}")
print(f"总保费投入: {total_premium:,.0f}元")

# 获取提取信息
extract_data = {}
for i, row in df.iterrows():
    year = int(row['保单年度'])
    extract_amount = row['提取金额'] if pd.notna(row['提取金额']) else 0
    if extract_amount > 0:
        extract_data[year] = float(extract_amount)

print(f"\n提取年份: {len(extract_data)}个")
for year, amount in sorted(extract_data.items()):
    print(f"  第{year}年: {amount:,.0f}元")

# 计算IRR结果
print(f"\n=== IRR计算结果（所有数据行）===")
print(f"{'年份':>4} | {'提取金额（元）':>12} | {'已提取金额累计':>12} | {'提取后总价值':>12} | {'现金流期数':>8} | {'IRR%':>8}")
print("-" * 75)

all_results = []
negative_count = 0
cumulative_extract = 0

for i, row in df.iterrows():
    year = int(row['保单年度'])
    final_value = row['提取后总价值']
    
    if isinstance(final_value, str):
        final_value = float(final_value.replace(',', ''))
    
    if pd.notna(final_value) and final_value > 0:
        # 构建现金流
        cash_flows = premiums.copy()
        
        current_extract = extract_data.get(year, 0)
        
        # 计算累计提取金额（截止到当前年份）
        cumulative_extract = sum(extract_data.get(y, 0) for y in range(1, year + 1))
        
        # 第6年到第year年期间的现金流
        for y in range(pay_years + 1, year + 1):
            if y in extract_data:
                cash_flows.append(extract_data[y])
            else:
                cash_flows.append(0)
        
        cash_flows.append(final_value)
        periods = len(cash_flows)
        
        # 计算IRR
        try:
            irr = calculate_irr(cash_flows)
            irr_pct = irr * 100
            
            # 显示所有行，但负数IRR显示为空
            if irr_pct > 0:
                irr_display = f"{irr_pct:6.2f}%"
            else:
                irr_display = "      -"
                negative_count += 1
            
            print(f"{year:>4} | {current_extract:>12,.0f} | {cumulative_extract:>12,.0f} | {final_value:>12,.0f} | {periods:>8} | {irr_display:>8}")
            
            all_results.append({
                'year': year,
                'extract_amount': current_extract,
                'cumulative_extract': cumulative_extract,
                'final_value': final_value,
                'periods': periods,
                'irr_pct': irr_pct if irr_pct > 0 else None,
                'irr_display': irr_display
            })
            
        except:
            # 计算失败也显示行，IRR为空
            irr_display = "      -"
            negative_count += 1
            print(f"{year:>4} | {current_extract:>12,.0f} | {cumulative_extract:>12,.0f} | {final_value:>12,.0f} | {periods:>8} | {irr_display:>8}")
            
            all_results.append({
                'year': year,
                'extract_amount': current_extract,
                'cumulative_extract': cumulative_extract,
                'final_value': final_value,
                'periods': periods,
                'irr_pct': None,
                'irr_display': irr_display
            })

print()
print(f"总计显示: {len(all_results)} 行数据")
print(f"其中负数或无效IRR: {negative_count} 个（显示为空）")

# 显示关键年份汇总（只显示有效IRR）
print(f"\n=== 关键年份IRR汇总（仅正数）===")
key_years = [5, 10, 15, 20, 25, 30, 35]
for year in key_years:
    result = next((r for r in all_results if r['year'] == year and r['irr_pct'] is not None), None)
    if result:
        extract_info = f"(提取: {result['extract_amount']:,.0f}元)" if result['extract_amount'] > 0 else ""
        print(f"第{year:2d}年: {result['irr_pct']:6.2f}% {extract_info}")

# 保存结果到Excel
try:
    # 创建结果数据
    result_data = []
    for result in all_results:
        irr_value = f"{result['irr_pct']:.2f}%" if result['irr_pct'] is not None else ""
        result_data.append([
            f"第{result['year']}年",
            result['extract_amount'],
            result['cumulative_extract'],
            result['final_value'],
            result['periods'],
            irr_value
        ])
    
    # 创建DataFrame
    result_df = pd.DataFrame(result_data, columns=[
        '年份', '提取金额（元）', '已提取金额累计（元）', '提取后总价值（元）', '现金流期数', 'IRR(%)'
    ])
    
    # 保存到Excel
    with pd.ExcelWriter('IRR计算结果_最终版.xlsx', engine='openpyxl') as writer:
        # 写入说明信息
        info_data = [
            ['IRR计算结果 - 最终版'],
            [''],
            ['说明: 显示所有数据行，IRR为负数时显示为空'],
            [f'负数或无效IRR结果: {negative_count}个'],
            [''],
            ['保费信息'],
            [f'保费现金流: {premiums}'],
            [f'总保费投入: {total_premium:,.0f}元'],
            [f'提取年份: {len(extract_data)}个'],
            ['']
        ]
        
        info_df = pd.DataFrame(info_data)
        info_df.to_excel(writer, sheet_name='计算结果', index=False, header=False, startrow=0)
        
        # 写入结果数据
        result_df.to_excel(writer, sheet_name='计算结果', index=False, startrow=len(info_data))
    
    print(f"\n✅ 结果已保存至: IRR计算结果_最终版.xlsx")
    
except Exception as e:
    print(f"\n❌ 保存Excel失败: {e}")

print(f"\n=== 计算完成 ===")
print(f"✅ 保费现金流: {premiums}")
print(f"✅ 总保费投入: {total_premium:,.0f}元")
print(f"✅ 提取年份: {len(extract_data)}个")
print(f"✅ 显示所有数据: {len(all_results)}行")
print(f"✅ 负数IRR显示为空: {negative_count}个")
print(f"✅ IRR保留2位小数，增加累计提取金额列")
print(f"✅ 最终版IRR计算完成！") 