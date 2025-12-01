#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
最终优化版图表生成脚本
修复问题：
1. 星级分布图文字重叠
2. 产品优缺点图字体显示问题
3. 移除不必要的标题说明
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

class FinalChartGenerator:
    """最终版图表生成器"""

    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None
        self.product_insights = {}

    def load_data(self):
        """加载数据"""
        print("=" * 80)
        print("📊 加载数据")
        print("=" * 80)
        self.df = pd.read_excel(self.file_path)
        print(f"✓ 成功读取 {len(self.df)} 条评论数据\n")

    def analyze_product_pros_cons(self):
        """分析产品优缺点"""
        print("🔍 分析产品优缺点...")

        product_col = 'Product Category'
        rating_col = 'Rating'
        review_col = 'Review'

        positive_keywords = {
            'quality': '质量优秀', 'great': '表现优异', 'excellent': '卓越性能',
            'perfect': '完美体验', 'fast': '速度快', 'amazing': '令人惊艳',
            'love': '用户喜爱', 'best': '最佳选择', 'good': '良好',
            'worth': '物有所值', 'value': '高性价比', 'recommend': '推荐',
            'powerful': '功能强大', 'compact': '小巧便携', 'durable': '耐用',
            'comfortable': '舒适', 'clear': '清晰', 'easy': '易用'
        }

        negative_keywords = {
            'bad': '质量差', 'poor': '表现不佳', 'disappointed': '令人失望',
            'waste': '浪费', 'broke': '损坏', 'stopped': '停止工作',
            'problem': '存在问题', 'issue': '有缺陷', 'weak': '性能弱',
            'lost': '迷失/丢失', 'fail': '失败', 'not work': '不工作',
            'stuck': '卡住', 'slow': '速度慢', 'complaint': '投诉',
            'difficult': '困难', 'complicated': '复杂', 'delay': '延迟'
        }

        products = self.df[product_col].unique()

        for product in products:
            product_df = self.df[self.df[product_col] == product]
            avg_rating = product_df[rating_col].mean()
            review_count = len(product_df)

            good_reviews = product_df[product_df[rating_col] >= 4][review_col].dropna()
            bad_reviews = product_df[product_df[rating_col] <= 2][review_col].dropna()

            pros = []
            for keyword, desc in positive_keywords.items():
                count = sum(good_reviews.str.lower().str.contains(keyword, na=False))
                if count > 0:
                    pros.append((desc, count, keyword))
            pros.sort(key=lambda x: x[1], reverse=True)

            cons = []
            for keyword, desc in negative_keywords.items():
                count = sum(bad_reviews.str.lower().str.contains(keyword, na=False))
                if count > 0:
                    cons.append((desc, count, keyword))
            cons.sort(key=lambda x: x[1], reverse=True)

            self.product_insights[product] = {
                'avg_rating': avg_rating,
                'review_count': review_count,
                'pros': pros[:5],
                'cons': cons[:3],
                'good_review_rate': len(good_reviews) / review_count * 100 if review_count > 0 else 0
            }

        print(f"✓ 完成 {len(products)} 个产品的优缺点分析\n")

    def plot_rating_distribution(self):
        """优化的星级分布图"""
        print("📊 生成星级分布图...")

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
        fig.suptitle('星级分布图', fontsize=18, fontweight='bold',
                    fontfamily='Microsoft YaHei')

        rating_col = 'Rating'

        # 左图：柱状图 - 优化标签位置避免重叠
        rating_counts = self.df[rating_col].value_counts().sort_index()
        colors_bar = plt.cm.Blues(np.linspace(0.4, 0.9, len(rating_counts)))

        bars = ax1.bar(rating_counts.index, rating_counts.values,
                      color=colors_bar, edgecolor='black', linewidth=0.5, width=0.15)

        ax1.set_xlabel('用户评分(星级)', fontsize=12, fontfamily='Microsoft YaHei')
        ax1.set_ylabel('评论数量', fontsize=12, fontfamily='Microsoft YaHei')
        ax1.set_title('各星级评论数量分布', fontsize=14, fontfamily='Microsoft YaHei')
        ax1.grid(axis='y', alpha=0.3)
        ax1.set_xlim(0.5, 5.5)

        # 优化数值标签 - 放在柱子上方，增加间距
        for bar, (rating, count) in zip(bars, rating_counts.items()):
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height + 3,
                    f'{int(count)}',
                    ha='center', va='bottom', fontsize=10,
                    fontweight='bold', fontfamily='Microsoft YaHei')

        # 右图：饼图
        rating_bins = pd.cut(self.df[rating_col], bins=[0, 1, 2, 3, 4, 5],
                            labels=['1星', '2星', '3星', '4星', '5星'], include_lowest=True)
        rating_grouped = rating_bins.value_counts().sort_index()

        colors_pie = ['#ff4444', '#ff9944', '#ffdd44', '#88dd44', '#44dd44']
        explode = [0.05 if i == rating_grouped.values.argmax() else 0
                  for i in range(len(rating_grouped))]

        wedges, texts, autotexts = ax2.pie(
            rating_grouped.values,
            labels=rating_grouped.index,
            autopct='%1.1f%%',
            colors=colors_pie,
            explode=explode,
            startangle=90,
            textprops={'fontsize': 11, 'fontfamily': 'Microsoft YaHei'}
        )

        ax2.set_title('星级占比分布', fontsize=14, fontfamily='Microsoft YaHei')

        # 优化饼图文字
        for text in texts:
            text.set_fontfamily('Microsoft YaHei')
            text.set_fontsize(11)

        for autotext in autotexts:
            autotext.set_fontfamily('Microsoft YaHei')
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(10)

        plt.tight_layout()
        plt.savefig('rating_distribution_improved.png', dpi=300, bbox_inches='tight')
        print("✓ 已生成: rating_distribution_improved.png\n")
        plt.close()

    def plot_all_products_details(self):
        """优化的产品优缺点详细图"""
        print("📊 生成产品优缺点详细图...")

        product_col = 'Product Category'
        rating_col = 'Rating'

        product_stats = self.df.groupby(product_col).agg({
            rating_col: ['mean', 'count']
        }).reset_index()
        product_stats.columns = ['Product', 'AvgRating', 'Count']
        product_stats = product_stats.sort_values('AvgRating', ascending=False)

        num_products = len(product_stats)

        # 创建大图表
        fig = plt.figure(figsize=(20, 4 * num_products))
        fig.suptitle('各产品优缺点详细分析', fontsize=22, fontweight='bold',
                    fontfamily='Microsoft YaHei', y=0.995)

        for idx, (_, row) in enumerate(product_stats.iterrows()):
            product = row['Product']

            if product not in self.product_insights:
                continue

            insight = self.product_insights[product]

            ax = plt.subplot(num_products, 1, idx + 1)
            ax.axis('off')

            # 背景色
            bg_color = plt.cm.Pastel1(idx % 9)
            rect = plt.Rectangle((0, 0), 1, 1, transform=ax.transAxes,
                                facecolor=bg_color, alpha=0.3, zorder=0)
            ax.add_patch(rect)

            # 产品标题
            title_y = 0.85
            ax.text(0.5, title_y, f"{product}", ha='center', fontsize=18,
                   fontweight='bold', transform=ax.transAxes,
                   fontfamily='Microsoft YaHei',
                   bbox=dict(boxstyle='round,pad=0.8', facecolor='white',
                           edgecolor='black', linewidth=2))

            # 基础数据
            info_y = 0.7
            info_text = f"⭐ 平均评分: {insight['avg_rating']:.2f}★  |  📝 评论数: {insight['review_count']}条  |  👍 好评率: {insight['good_review_rate']:.1f}%"
            ax.text(0.5, info_y, info_text, ha='center', fontsize=13,
                   transform=ax.transAxes, fontfamily='Microsoft YaHei',
                   bbox=dict(boxstyle='round,pad=0.5', facecolor='lightyellow', alpha=0.8))

            # 优点部分 (左侧)
            pros_y = 0.55
            ax.text(0.05, pros_y, "✅ 主要优点", fontsize=14, fontweight='bold',
                   transform=ax.transAxes, color='darkgreen',
                   fontfamily='Microsoft YaHei')

            pros_content_y = 0.45
            if insight['pros']:
                pros_text = ""
                for i, (desc, count, kw) in enumerate(insight['pros'][:5], 1):
                    pros_text += f"{i}. {desc} (提及{count}次)\n"
                ax.text(0.08, pros_content_y, pros_text, fontsize=12,
                       transform=ax.transAxes, fontfamily='Microsoft YaHei',
                       verticalalignment='top',
                       bbox=dict(boxstyle='round,pad=0.8', facecolor='lightgreen',
                               alpha=0.3, edgecolor='green', linewidth=1))
            else:
                ax.text(0.08, pros_content_y, "暂无明显优点关键词", fontsize=11,
                       transform=ax.transAxes, fontfamily='Microsoft YaHei',
                       style='italic')

            # 缺点部分 (右侧)
            cons_y = 0.55
            ax.text(0.55, cons_y, "⚠️ 主要缺点", fontsize=14, fontweight='bold',
                   transform=ax.transAxes, color='darkred',
                   fontfamily='Microsoft YaHei')

            cons_content_y = 0.45
            if insight['cons']:
                cons_text = ""
                for i, (desc, count, kw) in enumerate(insight['cons'][:5], 1):
                    cons_text += f"{i}. {desc} (提及{count}次)\n"
                ax.text(0.58, cons_content_y, cons_text, fontsize=12,
                       transform=ax.transAxes, fontfamily='Microsoft YaHei',
                       verticalalignment='top',
                       bbox=dict(boxstyle='round,pad=0.8', facecolor='#ffcccc',
                               alpha=0.3, edgecolor='red', linewidth=1))
            else:
                ax.text(0.58, cons_content_y, "✨ 暂无明显问题，表现优秀！",
                       fontsize=11, transform=ax.transAxes,
                       fontfamily='Microsoft YaHei',
                       style='italic', color='green', fontweight='bold')

            # 添加分隔线
            if idx < num_products - 1:
                ax.plot([0, 1], [0.05, 0.05], 'k-', linewidth=2, transform=ax.transAxes)

        plt.tight_layout()
        plt.savefig('all_products_details.png', dpi=300, bbox_inches='tight')
        print("✓ 已生成: all_products_details.png\n")
        plt.close()

    def run(self):
        """运行完整流程"""
        self.load_data()
        self.analyze_product_pros_cons()
        self.plot_rating_distribution()
        self.plot_all_products_details()

        print("=" * 80)
        print("✅ 所有图表生成完成！")
        print("=" * 80)

if __name__ == "__main__":
    generator = FinalChartGenerator(r"e:\Data\VS\AI\PP\PP\reviews_data.xlsx")
    generator.run()
