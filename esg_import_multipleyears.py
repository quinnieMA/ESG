# -*- coding: utf-8 -*-
"""
Created on Tue Jun 10 12:57:09 2025

@author: FM
"""

import os
from WindPy import *
w.start()

from global_options_cn import DataPaths, BASE_DIR
from pathlib import Path
# 设置路径
#FOLDER =BASE_DIR

# 定义文件路径
code_file = BASE_DIR/"code.txt"
#output_folder = os.path.join(BASE_DIR,'esg')
output_folder = Path(r"C:\Users\FM\OneDrive\NLP1\DATA\esg")


# 从文件读取股票代码
with open(code_file, "r", encoding="utf-8") as f:
    stock_codes = f.read().strip()  # 直接读取整个字符串，已经是逗号分隔格式

# 检查代码是否有效
if not stock_codes:
    raise ValueError("股票代码文件为空或格式不正确")

print("从文件读取的股票代码字符串:")
print(stock_codes)

# Wind API 查询字段
esg_fields = "esg_conclusion,esg_rating_wind,esg_score_wind,esg_mgmtscore_wind," \
             "esg_eventscore_wind,esg_escore_wind,esg_sscore_wind,esg_gscore_wind," \
             "esg_rating_wind2,esg_score_wind2,esg_mgmtscore_wind2,esg_eventscore_wind2," \
             "esg_escore_wind2,esg_sscore_wind2,esg_gscore_wind2,esg_ghg1_wind," \
             "esg_ghgr1_wind,esg_ghg2_wind,esg_ghgr2_wind,esg_ghg12_wind,esg_ghgr12_wind," \
             "esg_em_wind,esg_cr_wind,esg_rating_ssi,esg_rating_ftserussell,esg_rating"

# 年份遍历（2016到2024）
for year in range(2025, 2015, -1):  # 
    print(f"\n正在处理 {year} 年数据...")
    
    # 构建查询参数
    query_params = f"tradeDate={year}0630;rptDate={year}1231"
    
    # 执行查询
    data = w.wss(stock_codes, esg_fields, query_params, usedf=True)
    
    # 处理结果
    if isinstance(data, tuple) and len(data) == 2:
        error_code, result_df = data
        if error_code == 0:
            # 添加年份列
            result_df['year'] = year
            
            # 构建带年份的文件名
            output_file = output_folder / f"wind_esg_results_{year}.xlsx"
            csv_file = output_folder / f"wind_esg_results_{year}.csv"
            
            try:
                # 保存Excel
                result_df.to_excel(output_file, index=True)
                print(f"{year}年数据已保存至: {output_file}")
                
                # 保存CSV
                result_df.to_csv(csv_file, encoding='utf_8_sig', index=True)
                print(f"{year}年CSV备份已保存至: {csv_file}")
                
                # 显示前5行结果
                print(f"{year}年数据预览:")
                print(result_df.head())
                
            except Exception as e:
                print(f"{year}年数据保存失败: {e}")
        else:
            print(f"{year}年查询失败，错误代码: {error_code}")
    else:
        print(f"{year}年返回数据格式异常")

print("\n所有年份数据处理完成！")