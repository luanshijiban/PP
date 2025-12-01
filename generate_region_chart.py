#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
生成区域分析图 - 使用微软雅黑字体
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import sys
import io
import warnings

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
warnings.filterwarnings('ignore')

# 设置中文字体为微软雅黑
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['figure.dpi'] = 100

sns.set_style("whitegrid")

def generate_region_analysis():
    """生成区域分析图"""
    print("=" * 80)
    print("📊 生成区域分析图")
    print("=" * 80)

    # 读取数据
    df = pd.read_excel(r"e:\Data\VS\AI\PP\PP\reviews_data.xlsx")
    print(f"✓ 成功读取 {len(df)} 条评论数据\n")

    # 按国家统计
    region_stats = df.groupby('Country').agg({
        'Rating': ['mean', 'count']
    }).reset_index()
    region_stats.columns = ['Country', 'AvgRating', 'Count']
    region_stats = region_stats.sort_values('Count', ascending=False)

    # 创建图表
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
    fig.suptitle('区域分析图', fontsize=18, fontweight='bold',
                fontfamily='Microsoft YaHei')

    # 左图：评论数量TOP10
    top10_count = region_stats.head(10)
    colors1 = plt.cm.Blues(np.linspace(0.4, 0.9, len(top10_count)))

    bars1 = ax1.barh(top10_count['Country'], top10_count['Count'],
                     color=colors1, edgecolor='black', linewidth=0.5)
    ax1.set_xlabel('评论数量', fontsize=12, fontfamily='Microsoft YaHei')
    ax1.set_title('各国家/地区评论数量 TOP10', fontsize=14, fontfamily='Microsoft YaHei')
    ax1.grid(axis='x', alpha=0.3)
    ax1.invert_yaxis()

    # 添加数值标签
    for i, (idx, row) in enumerate(top10_count.iterrows()):
        ax1.text(row['Count'] + 0.5, i, f"{int(row['Count'])}",
                va='center', fontsize=9, fontfamily='Microsoft YaHei')

    # 右图：平均评分TOP10
    top10_rating = region_stats.sort_values('AvgRating', ascending=False).head(10)
    colors2 = plt.cm.RdYlGn(np.linspace(0.3, 0.9, len(top10_rating)))

    bars2 = ax2.barh(top10_rating['Country'], top10_rating['AvgRating'],
                     color=colors2, edgecolor='black', linewidth=0.5)
    ax2.set_xlabel('平均评分(星)', fontsize=12, fontfamily='Microsoft YaHei')
    ax2.set_title('各国家/地区平均评分 TOP10', fontsize=14, fontfamily='Microsoft YaHei')
    ax2.set_xlim(0, 5.5)
    ax2.grid(axis='x', alpha=0.3)
    ax2.invert_yaxis()

    # 添加数值标签
    for i, (idx, row) in enumerate(top10_rating.iterrows()):
        ax2.text(row['AvgRating'] + 0.05, i, f"{row['AvgRating']:.2f}★",
                va='center', fontsize=9, fontfamily='Microsoft YaHei')

    plt.tight_layout()
    plt.savefig('region_analysis.png', dpi=300, bbox_inches='tight')
    print("✓ 已生成: region_analysis.png\n")
    plt.close()

    print("=" * 80)
    print("✅ 区域分析图生成完成！")
    print("=" * 80)

if __name__ == "__main__":
    generate_region_analysis()
