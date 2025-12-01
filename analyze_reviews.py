#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
用户评论数据分析脚本
完成从数据读取、清洗、分析到可视化的完整流程
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import warnings
import json
import sys
import os
import io

# 设置输出编码为UTF-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

warnings.filterwarnings('ignore')

# 设置中文字体支持
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

# 设置绘图风格
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (12, 8)

class ReviewAnalyzer:
    """用户评论数据分析器"""

    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None
        self.stats = {}
        self.insights = []

    def load_data(self):
        """加载Excel数据"""
        print("=" * 60)
        print("第一步：读取数据文件")
        print("=" * 60)
        try:
            self.df = pd.read_excel(self.file_path)
            print(f"✓ 成功读取数据: {len(self.df)} 条记录")
            print(f"\n数据列名: {list(self.df.columns)}")
            print(f"\n数据前5行预览:")
            print(self.df.head())
            print(f"\n数据基本信息:")
            print(self.df.info())
            return True
        except Exception as e:
            print(f"✗ 读取数据失败: {e}")
            return False

    def clean_data(self):
        """数据清洗与预处理"""
        print("\n" + "=" * 60)
        print("第二步：数据清洗与预处理")
        print("=" * 60)

        initial_count = len(self.df)
        print(f"初始数据量: {initial_count} 条")

        # 显示缺失值情况
        missing_summary = self.df.isnull().sum()
        print(f"\n缺失值统计:")
        print(missing_summary[missing_summary > 0])

        # 数据类型检查和转换
        print(f"\n数据类型检查:")
        for col in self.df.columns:
            print(f"  {col}: {self.df[col].dtype}")

        # 尝试识别日期列并转换
        for col in self.df.columns:
            if 'date' in col.lower() or 'time' in col.lower() or '日期' in col or '时间' in col:
                try:
                    self.df[col] = pd.to_datetime(self.df[col])
                    print(f"✓ 成功转换日期列: {col}")
                except:
                    pass

        # 识别星级/评分列
        rating_cols = [col for col in self.df.columns if 'rating' in col.lower() or
                      'star' in col.lower() or '星' in col or '评分' in col or 'score' in col.lower()]

        if rating_cols:
            print(f"\n识别到评分列: {rating_cols}")
            for col in rating_cols:
                print(f"\n{col} 统计:")
                print(self.df[col].describe())

        print(f"\n清洗后数据量: {len(self.df)} 条")
        print(f"数据完整性: {(1 - self.df.isnull().sum().sum() / (len(self.df) * len(self.df.columns))) * 100:.2f}%")

        return True

    def analyze_ratings(self):
        """整体星级分析"""
        print("\n" + "=" * 60)
        print("第三步：整体星级分析")
        print("=" * 60)

        # 查找评分列
        rating_col = None
        for col in self.df.columns:
            if 'rating' in col.lower() or 'star' in col.lower() or '星' in col or '评分' in col:
                rating_col = col
                break

        if rating_col:
            ratings = self.df[rating_col].dropna()

            # 基础统计
            print(f"\n【星级分布统计】")
            print(f"总评论数: {len(ratings)}")
            print(f"平均评分: {ratings.mean():.2f}")
            print(f"中位数评分: {ratings.median():.2f}")
            print(f"最高评分: {ratings.max()}")
            print(f"最低评分: {ratings.min()}")
            print(f"标准差: {ratings.std():.2f}")

            # 星级分布
            print(f"\n【各星级数量与占比】")
            rating_dist = ratings.value_counts().sort_index(ascending=False)
            for star, count in rating_dist.items():
                pct = count / len(ratings) * 100
                print(f"{star}星: {count} 条 ({pct:.1f}%)")

            # 好评率计算（4-5星为好评）
            if ratings.max() >= 4:
                good_reviews = ratings[ratings >= 4].count()
                good_rate = good_reviews / len(ratings) * 100
                print(f"\n好评率 (4-5星): {good_rate:.1f}%")

            # 保存统计数据
            self.stats['rating_stats'] = {
                'mean': float(ratings.mean()),
                'median': float(ratings.median()),
                'std': float(ratings.std()),
                'distribution': rating_dist.to_dict()
            }

            # 生成星级分布图
            self.plot_rating_distribution(ratings, rating_col)

        else:
            print("⚠ 未找到评分列")

    def analyze_trends(self):
        """消费者趋势分析"""
        print("\n" + "=" * 60)
        print("第四步：消费者趋势分析（时间序列）")
        print("=" * 60)

        # 查找日期列
        date_col = None
        for col in self.df.columns:
            if self.df[col].dtype == 'datetime64[ns]':
                date_col = col
                break

        if date_col:
            print(f"使用日期列: {date_col}")

            # 按月统计
            self.df['year_month'] = self.df[date_col].dt.to_period('M')
            monthly_stats = self.df.groupby('year_month').size()

            print(f"\n【月度评论趋势】")
            for period, count in monthly_stats.items():
                print(f"{period}: {count} 条")

            # 查找评分列用于趋势分析
            rating_col = None
            for col in self.df.columns:
                if 'rating' in col.lower() or 'star' in col.lower() or '星' in col:
                    rating_col = col
                    break

            if rating_col:
                monthly_rating = self.df.groupby('year_month')[rating_col].mean()
                print(f"\n【月度平均评分趋势】")
                for period, rating in monthly_rating.items():
                    print(f"{period}: {rating:.2f}星")

                # 生成趋势图
                self.plot_trends(monthly_stats, monthly_rating, date_col)
        else:
            print("⚠ 未找到日期列，无法进行时间趋势分析")

    def analyze_products(self):
        """产品维度分析"""
        print("\n" + "=" * 60)
        print("第五步：产品维度深度分析")
        print("=" * 60)

        # 查找产品相关列
        product_cols = [col for col in self.df.columns if 'product' in col.lower() or
                       '产品' in col or 'item' in col.lower() or 'sku' in col.lower()]

        if product_cols:
            product_col = product_cols[0]
            print(f"使用产品列: {product_col}")

            # 产品数量统计
            product_counts = self.df[product_col].value_counts()
            print(f"\n【产品评论数量TOP10】")
            for i, (product, count) in enumerate(product_counts.head(10).items(), 1):
                pct = count / len(self.df) * 100
                print(f"{i}. {product}: {count} 条 ({pct:.1f}%)")

            # 如果有评分，按产品统计平均评分
            rating_col = None
            for col in self.df.columns:
                if 'rating' in col.lower() or 'star' in col.lower() or '星' in col:
                    rating_col = col
                    break

            if rating_col:
                product_ratings = self.df.groupby(product_col)[rating_col].agg(['mean', 'count'])
                product_ratings = product_ratings.sort_values('mean', ascending=False)

                print(f"\n【产品平均评分TOP10】")
                for i, (product, row) in enumerate(product_ratings.head(10).iterrows(), 1):
                    print(f"{i}. {product}: {row['mean']:.2f}星 (基于{int(row['count'])}条评论)")

                print(f"\n【产品平均评分BOTTOM10】")
                for i, (product, row) in enumerate(product_ratings.tail(10).iterrows(), 1):
                    print(f"{i}. {product}: {row['mean']:.2f}星 (基于{int(row['count'])}条评论)")

                # 生成产品分析图
                self.plot_product_analysis(product_counts, product_ratings, product_col)

        else:
            print("⚠ 未找到产品列")

    def analyze_regions(self):
        """区域分析"""
        print("\n" + "=" * 60)
        print("第六步:区域分布与差异分析")
        print("=" * 60)

        # 查找区域相关列
        region_cols = [col for col in self.df.columns if 'region' in col.lower() or
                      'country' in col.lower() or 'location' in col.lower() or
                      '地区' in col or '国家' in col or '区域' in col]

        if region_cols:
            region_col = region_cols[0]
            print(f"使用区域列: {region_col}")

            # 区域分布统计
            region_counts = self.df[region_col].value_counts()
            print(f"\n【区域评论数量分布TOP10】")
            for i, (region, count) in enumerate(region_counts.head(10).items(), 1):
                pct = count / len(self.df) * 100
                print(f"{i}. {region}: {count} 条 ({pct:.1f}%)")

            # 如果有评分，按区域统计
            rating_col = None
            for col in self.df.columns:
                if 'rating' in col.lower() or 'star' in col.lower() or '星' in col:
                    rating_col = col
                    break

            if rating_col:
                region_ratings = self.df.groupby(region_col)[rating_col].agg(['mean', 'count'])
                region_ratings = region_ratings.sort_values('mean', ascending=False)

                print(f"\n【区域平均评分TOP10】")
                for i, (region, row) in enumerate(region_ratings.head(10).iterrows(), 1):
                    print(f"{i}. {region}: {row['mean']:.2f}星 (基于{int(row['count'])}条评论)")

                # 生成区域分析图
                self.plot_region_analysis(region_counts, region_ratings, region_col)
        else:
            print("⚠ 未找到区域列")

    def detect_anomalies(self):
        """异常数据识别"""
        print("\n" + "=" * 60)
        print("第七步：异常数据识别与分析")
        print("=" * 60)

        anomalies = []

        # 1. 检查评分异常
        rating_col = None
        for col in self.df.columns:
            if 'rating' in col.lower() or 'star' in col.lower() or '星' in col:
                rating_col = col
                break

        if rating_col:
            # 极端评分的异常集中
            rating_counts = self.df[rating_col].value_counts()
            if rating_counts.max() / len(self.df) > 0.7:
                dominant_rating = rating_counts.idxmax()
                anomalies.append(f"⚠ 评分异常集中: {dominant_rating}星占比{rating_counts.max()/len(self.df)*100:.1f}%")

        # 2. 检查重复评论
        comment_cols = [col for col in self.df.columns if 'comment' in col.lower() or
                       'review' in col.lower() or '评论' in col or '内容' in col]

        if comment_cols:
            comment_col = comment_cols[0]
            duplicates = self.df[comment_col].duplicated().sum()
            if duplicates > 0:
                dup_rate = duplicates / len(self.df) * 100
                anomalies.append(f"⚠ 发现重复评论: {duplicates} 条 ({dup_rate:.1f}%)")

                # 显示重复最多的评论
                top_duplicates = self.df[comment_col].value_counts().head(5)
                if top_duplicates.iloc[0] > 5:
                    print(f"\n【重复次数最多的评论】")
                    for i, (comment, count) in enumerate(top_duplicates.items(), 1):
                        if count > 1:
                            print(f"{i}. 重复{count}次: {str(comment)[:50]}...")

        # 3. 检查时间异常
        date_cols = [col for col in self.df.columns if self.df[col].dtype == 'datetime64[ns]']
        if date_cols:
            date_col = date_cols[0]
            # 检查是否有未来日期
            future_dates = self.df[self.df[date_col] > datetime.now()]
            if len(future_dates) > 0:
                anomalies.append(f"⚠ 发现未来日期: {len(future_dates)} 条")

            # 检查评论爆发式增长
            if 'year_month' in self.df.columns:
                monthly_counts = self.df.groupby('year_month').size()
                if len(monthly_counts) > 1:
                    growth_rates = monthly_counts.pct_change()
                    extreme_growth = growth_rates[growth_rates > 2]  # 增长超过200%
                    if len(extreme_growth) > 0:
                        for period, rate in extreme_growth.items():
                            anomalies.append(f"⚠ 评论数激增: {period} 增长{rate*100:.0f}%")

        # 4. 检查数据完整性异常
        for col in self.df.columns:
            null_rate = self.df[col].isnull().sum() / len(self.df)
            if null_rate > 0.3:  # 缺失率超过30%
                anomalies.append(f"⚠ 字段 '{col}' 缺失率过高: {null_rate*100:.1f}%")

        # 输出异常发现
        if anomalies:
            print(f"\n【发现 {len(anomalies)} 个异常现象】")
            for i, anomaly in enumerate(anomalies, 1):
                print(f"{i}. {anomaly}")
        else:
            print("\n✓ 未发现明显异常")

        self.stats['anomalies'] = anomalies
        return anomalies

    def plot_rating_distribution(self, ratings, rating_col):
        """绘制星级分布图"""
        fig, axes = plt.subplots(1, 2, figsize=(15, 5))

        # 柱状图
        rating_counts = ratings.value_counts().sort_index()
        axes[0].bar(rating_counts.index, rating_counts.values, color='steelblue', alpha=0.7)
        axes[0].set_xlabel('评分（星级）', fontsize=12)
        axes[0].set_ylabel('数量', fontsize=12)
        axes[0].set_title('星级分布 - 柱状图', fontsize=14, fontweight='bold')
        axes[0].grid(axis='y', alpha=0.3)

        # 饼图
        colors = plt.cm.RdYlGn(np.linspace(0.2, 0.9, len(rating_counts)))
        axes[1].pie(rating_counts.values, labels=[f'{int(x)}星' for x in rating_counts.index],
                   autopct='%1.1f%%', colors=colors, startangle=90)
        axes[1].set_title('星级分布 - 占比图', fontsize=14, fontweight='bold')

        plt.tight_layout()
        plt.savefig('rating_distribution.png', dpi=300, bbox_inches='tight')
        print("\n✓ 已生成图表: rating_distribution.png")
        plt.close()

    def plot_trends(self, monthly_stats, monthly_rating, date_col):
        """绘制趋势图"""
        fig, axes = plt.subplots(2, 1, figsize=(14, 10))

        # 评论数量趋势
        x_labels = [str(p) for p in monthly_stats.index]
        axes[0].plot(x_labels, monthly_stats.values, marker='o', linewidth=2,
                    markersize=8, color='steelblue', label='评论数量')
        axes[0].fill_between(range(len(x_labels)), monthly_stats.values, alpha=0.3, color='steelblue')
        axes[0].set_xlabel('时间', fontsize=12)
        axes[0].set_ylabel('评论数量', fontsize=12)
        axes[0].set_title('月度评论数量趋势', fontsize=14, fontweight='bold')
        axes[0].legend()
        axes[0].grid(True, alpha=0.3)
        axes[0].tick_params(axis='x', rotation=45)

        # 平均评分趋势
        x_labels_rating = [str(p) for p in monthly_rating.index]
        axes[1].plot(x_labels_rating, monthly_rating.values, marker='s', linewidth=2,
                    markersize=8, color='coral', label='平均评分')
        axes[1].fill_between(range(len(x_labels_rating)), monthly_rating.values, alpha=0.3, color='coral')
        axes[1].set_xlabel('时间', fontsize=12)
        axes[1].set_ylabel('平均评分（星）', fontsize=12)
        axes[1].set_title('月度平均评分趋势', fontsize=14, fontweight='bold')
        axes[1].legend()
        axes[1].grid(True, alpha=0.3)
        axes[1].tick_params(axis='x', rotation=45)

        plt.tight_layout()
        plt.savefig('trends_analysis.png', dpi=300, bbox_inches='tight')
        print("\n✓ 已生成图表: trends_analysis.png")
        plt.close()

    def plot_product_analysis(self, product_counts, product_ratings, product_col):
        """绘制产品分析图"""
        fig, axes = plt.subplots(2, 1, figsize=(14, 12))

        # TOP10产品评论数
        top10_counts = product_counts.head(10)
        axes[0].barh(range(len(top10_counts)), top10_counts.values, color='teal', alpha=0.7)
        axes[0].set_yticks(range(len(top10_counts)))
        axes[0].set_yticklabels([str(p)[:30] for p in top10_counts.index])
        axes[0].set_xlabel('评论数量', fontsize=12)
        axes[0].set_title('产品评论数量 TOP10', fontsize=14, fontweight='bold')
        axes[0].grid(axis='x', alpha=0.3)

        # TOP10产品平均评分
        top10_ratings = product_ratings.nlargest(10, 'mean')
        colors = plt.cm.RdYlGn(np.linspace(0.3, 0.9, len(top10_ratings)))
        bars = axes[1].barh(range(len(top10_ratings)), top10_ratings['mean'].values, color=colors, alpha=0.7)
        axes[1].set_yticks(range(len(top10_ratings)))
        axes[1].set_yticklabels([str(p)[:30] for p in top10_ratings.index])
        axes[1].set_xlabel('平均评分（星）', fontsize=12)
        axes[1].set_title('产品平均评分 TOP10', fontsize=14, fontweight='bold')
        axes[1].set_xlim(0, 5.5)
        axes[1].grid(axis='x', alpha=0.3)

        # 在柱子上标注数值
        for i, (bar, val) in enumerate(zip(bars, top10_ratings['mean'].values)):
            axes[1].text(val + 0.05, bar.get_y() + bar.get_height()/2,
                        f'{val:.2f}', va='center', fontsize=10)

        plt.tight_layout()
        plt.savefig('product_analysis.png', dpi=300, bbox_inches='tight')
        print("\n✓ 已生成图表: product_analysis.png")
        plt.close()

    def plot_region_analysis(self, region_counts, region_ratings, region_col):
        """绘制区域分析图"""
        fig, axes = plt.subplots(1, 2, figsize=(16, 6))

        # TOP10区域评论数
        top10_regions = region_counts.head(10)
        axes[0].barh(range(len(top10_regions)), top10_regions.values, color='purple', alpha=0.7)
        axes[0].set_yticks(range(len(top10_regions)))
        axes[0].set_yticklabels(top10_regions.index)
        axes[0].set_xlabel('评论数量', fontsize=12)
        axes[0].set_title('区域评论数量 TOP10', fontsize=14, fontweight='bold')
        axes[0].grid(axis='x', alpha=0.3)

        # TOP10区域平均评分
        top10_ratings = region_ratings.nlargest(10, 'mean')
        colors = plt.cm.RdYlGn(np.linspace(0.3, 0.9, len(top10_ratings)))
        bars = axes[1].barh(range(len(top10_ratings)), top10_ratings['mean'].values, color=colors, alpha=0.7)
        axes[1].set_yticks(range(len(top10_ratings)))
        axes[1].set_yticklabels(top10_ratings.index)
        axes[1].set_xlabel('平均评分（星）', fontsize=12)
        axes[1].set_title('区域平均评分 TOP10', fontsize=14, fontweight='bold')
        axes[1].set_xlim(0, 5.5)
        axes[1].grid(axis='x', alpha=0.3)

        plt.tight_layout()
        plt.savefig('region_analysis.png', dpi=300, bbox_inches='tight')
        print("\n✓ 已生成图表: region_analysis.png")
        plt.close()

    def generate_insights(self):
        """生成关键洞察"""
        print("\n" + "=" * 60)
        print("生成关键洞察")
        print("=" * 60)

        insights = []

        # 基于评分的洞察
        if 'rating_stats' in self.stats:
            mean_rating = self.stats['rating_stats']['mean']
            if mean_rating >= 4.5:
                insights.append(f"✓ 整体用户满意度很高，平均评分达到{mean_rating:.2f}星")
            elif mean_rating >= 4.0:
                insights.append(f"→ 整体用户满意度良好，平均评分{mean_rating:.2f}星，仍有提升空间")
            elif mean_rating >= 3.0:
                insights.append(f"⚠ 整体用户满意度中等，平均评分仅{mean_rating:.2f}星，需要重点关注")
            else:
                insights.append(f"✗ 整体用户满意度较低，平均评分{mean_rating:.2f}星，需紧急改善")

            # 评分分布洞察
            dist = self.stats['rating_stats']['distribution']
            if 5 in dist and 1 in dist:
                polarization = (dist.get(5, 0) + dist.get(1, 0)) / sum(dist.values())
                if polarization > 0.6:
                    insights.append("⚠ 评分呈现两极分化，建议深入分析好评和差评的具体原因")

        # 异常洞察
        if 'anomalies' in self.stats and self.stats['anomalies']:
            insights.append(f"⚠ 发现{len(self.stats['anomalies'])}个异常现象，需要进一步调查")

        self.insights = insights

        print("\n【关键洞察】")
        for i, insight in enumerate(insights, 1):
            print(f"{i}. {insight}")

        return insights

    def generate_report(self):
        """生成完整分析报告"""
        print("\n" + "=" * 60)
        print("生成完整分析报告")
        print("=" * 60)

        report = []
        report.append("=" * 80)
        report.append("用户评论数据分析报告")
        report.append("=" * 80)
        report.append(f"\n报告生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report.append(f"数据文件: {self.file_path}")
        report.append(f"数据记录数: {len(self.df)}")

        # 一、执行摘要
        report.append("\n\n【一、执行摘要】")
        report.append("-" * 80)
        if self.insights:
            for insight in self.insights:
                report.append(f"  • {insight}")

        # 二、数据概况
        report.append("\n\n【二、数据概况】")
        report.append("-" * 80)
        report.append(f"  • 总评论数: {len(self.df)} 条")
        report.append(f"  • 数据字段: {', '.join(self.df.columns)}")
        report.append(f"  • 数据时间范围: {self.df.select_dtypes(include=['datetime64']).min().min() if len(self.df.select_dtypes(include=['datetime64']).columns) > 0 else 'N/A'} 至 {self.df.select_dtypes(include=['datetime64']).max().max() if len(self.df.select_dtypes(include=['datetime64']).columns) > 0 else 'N/A'}")

        # 三、星级分析
        if 'rating_stats' in self.stats:
            report.append("\n\n【三、星级分析】")
            report.append("-" * 80)
            stats = self.stats['rating_stats']
            report.append(f"  • 平均评分: {stats['mean']:.2f} 星")
            report.append(f"  • 中位数评分: {stats['median']:.2f} 星")
            report.append(f"  • 标准差: {stats['std']:.2f}")
            report.append("\n  评分分布:")
            for star, count in sorted(stats['distribution'].items(), reverse=True):
                pct = count / sum(stats['distribution'].values()) * 100
                report.append(f"    {int(star)}星: {count} 条 ({pct:.1f}%)")

        # 四、改进建议
        report.append("\n\n【四、改进建议与业务策略】")
        report.append("-" * 80)

        # 基于分析生成建议
        recommendations = []

        if 'rating_stats' in self.stats:
            mean_rating = self.stats['rating_stats']['mean']
            dist = self.stats['rating_stats']['distribution']

            # 低分问题
            low_ratings = dist.get(1, 0) + dist.get(2, 0)
            if low_ratings > 0:
                low_rate = low_ratings / sum(dist.values()) * 100
                if low_rate > 20:
                    recommendations.append(f"【紧急】低分评价占比{low_rate:.1f}%，建议:\n"
                                         "    - 建立差评预警机制\n"
                                         "    - 对低分评论进行分类分析\n"
                                         "    - 优先解决高频负面问题")

            # 好评维护
            high_ratings = dist.get(4, 0) + dist.get(5, 0)
            if high_ratings > 0:
                high_rate = high_ratings / sum(dist.values()) * 100
                if high_rate > 60:
                    recommendations.append(f"【优势】好评率达{high_rate:.1f}%，建议:\n"
                                         "    - 总结高分评论的共性特征\n"
                                         "    - 将优势特性作为产品卖点\n"
                                         "    - 鼓励满意用户分享和推荐")

        # 异常相关建议
        if 'anomalies' in self.stats and self.stats['anomalies']:
            recommendations.append("【数据质量】发现异常数据，建议:\n"
                                 "    - 审查数据收集流程\n"
                                 "    - 建立数据质量监控\n"
                                 "    - 清洗和标准化历史数据")

        # 通用建议
        recommendations.append("【持续优化】通用建议:\n"
                             "    - 建立定期评论分析机制（周/月报）\n"
                             "    - 实施NPS（净推荐值）跟踪\n"
                             "    - 开展用户访谈深入了解痛点\n"
                             "    - 建立产品改进优先级矩阵")

        for i, rec in enumerate(recommendations, 1):
            report.append(f"\n  {i}. {rec}")

        # 五、附录
        report.append("\n\n【五、附录】")
        report.append("-" * 80)
        report.append("  生成的图表文件:")
        report.append("    • rating_distribution.png - 星级分布图")
        report.append("    • trends_analysis.png - 趋势分析图")
        report.append("    • product_analysis.png - 产品分析图")
        report.append("    • region_analysis.png - 区域分析图")

        report.append("\n" + "=" * 80)
        report.append("报告结束")
        report.append("=" * 80)

        # 保存报告
        report_text = '\n'.join(report)
        with open('analysis_report.txt', 'w', encoding='utf-8') as f:
            f.write(report_text)

        print("\n✓ 分析报告已保存: analysis_report.txt")
        print("\n" + report_text)

        return report_text

    def run_full_analysis(self):
        """运行完整分析流程"""
        print("\n" + "=" * 80)
        print("开始完整数据分析流程")
        print("=" * 80 + "\n")

        # 1. 加载数据
        if not self.load_data():
            return False

        # 2. 数据清洗
        self.clean_data()

        # 3. 星级分析
        self.analyze_ratings()

        # 4. 趋势分析
        self.analyze_trends()

        # 5. 产品分析
        self.analyze_products()

        # 6. 区域分析
        self.analyze_regions()

        # 7. 异常检测
        self.detect_anomalies()

        # 8. 生成洞察
        self.generate_insights()

        # 9. 生成报告
        self.generate_report()

        print("\n" + "=" * 80)
        print("✓ 分析完成！")
        print("=" * 80)

        return True


if __name__ == "__main__":
    # 设置文件路径
    file_path = r"d:\AI\AI 2.0\PP\reviews_data.xlsx"

    # 创建分析器并运行
    analyzer = ReviewAnalyzer(file_path)
    analyzer.run_full_analysis()
