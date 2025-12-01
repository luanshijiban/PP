#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
最终版用户评论数据分析脚本
- 修复中文字体显示问题（使用微软雅黑）
- 完整展示所有产品的优缺点分析
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import warnings
import sys
import io

# 设置输出编码为UTF-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
warnings.filterwarnings('ignore')

# 强制设置中文字体为微软雅黑
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['figure.dpi'] = 100

# 设置绘图风格
sns.set_style("whitegrid")

class FinalReviewAnalyzer:
    """最终版评论分析器"""

    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None
        self.product_insights = {}

    def load_data(self):
        """加载数据"""
        print("=" * 80)
        print("📊 开始数据分析")
        print("=" * 80)
        try:
            self.df = pd.read_excel(self.file_path)
            print(f"✓ 成功读取 {len(self.df)} 条评论数据\n")
            return True
        except Exception as e:
            print(f"✗ 读取失败: {e}")
            return False

    def analyze_product_pros_cons(self):
        """分析各产品的优缺点"""
        print("\n" + "=" * 80)
        print("🔍 产品优缺点深度分析")
        print("=" * 80)

        product_col = 'Product Category'
        rating_col = 'Rating'
        review_col = 'Review'

        products = self.df[product_col].unique()

        # 关键词定义
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

        for product in products:
            product_df = self.df[self.df[product_col] == product]

            avg_rating = product_df[rating_col].mean()
            review_count = len(product_df)

            good_reviews = product_df[product_df[rating_col] >= 4][review_col].dropna()
            bad_reviews = product_df[product_df[rating_col] <= 2][review_col].dropna()

            # 分析优点
            pros = []
            for keyword, desc in positive_keywords.items():
                count = sum(good_reviews.str.lower().str.contains(keyword, na=False))
                if count > 0:
                    pros.append((desc, count, keyword))
            pros.sort(key=lambda x: x[1], reverse=True)

            # 分析缺点
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

            # 打印分析结果
            print(f"\n【{product}】")
            print(f"  平均评分: {avg_rating:.2f}★ | 评论数: {review_count} | 好评率: {self.product_insights[product]['good_review_rate']:.1f}%")

            if pros:
                print(f"  ✅ 主要优点:")
                for i, (desc, count, kw) in enumerate(pros[:3], 1):
                    print(f"     {i}. {desc} (提及{count}次 - '{kw}')")
            else:
                print(f"  ✅ 主要优点: 暂无明显关键词")

            if cons:
                print(f"  ⚠️  主要缺点:")
                for i, (desc, count, kw) in enumerate(cons[:3], 1):
                    print(f"     {i}. {desc} (提及{count}次 - '{kw}')")
            else:
                print(f"  ⚠️  主要缺点: 暂无明显问题")

    def plot_improved_charts(self):
        """生成改进后的可视化图表"""
        print("\n" + "=" * 80)
        print("📈 生成改进版可视化图表")
        print("=" * 80)

        self.plot_rating_distribution_improved()
        self.plot_product_ranking_with_insights()
        self.plot_all_products_details()  # 新增：所有产品详细信息图

    def plot_rating_distribution_improved(self):
        """改进的星级分布图"""
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
        fig.suptitle('星级分布图', fontsize=18, fontweight='bold', fontname='Microsoft YaHei')

        rating_col = 'Rating'

        # 左图：柱状图
        rating_counts = self.df[rating_col].value_counts().sort_index()
        colors_bar = plt.cm.Blues(np.linspace(0.4, 0.9, len(rating_counts)))

        ax1.bar(rating_counts.index, rating_counts.values, color=colors_bar, edgecolor='black', linewidth=0.5)
        ax1.set_xlabel('用户评分(星级)', fontsize=12, fontname='Microsoft YaHei')
        ax1.set_ylabel('评论数量', fontsize=12, fontname='Microsoft YaHei')
        ax1.set_title('各星级评论数量分布', fontsize=14, fontname='Microsoft YaHei')
        ax1.grid(axis='y', alpha=0.3)

        for i, (rating, count) in enumerate(rating_counts.items()):
            ax1.text(rating, count, str(count), ha='center', va='bottom', fontsize=9, fontname='Microsoft YaHei')

        # 右图：饼图
        rating_bins = pd.cut(self.df[rating_col], bins=[0, 1, 2, 3, 4, 5],
                            labels=['1星', '2星', '3星', '4星', '5星'], include_lowest=True)
        rating_grouped = rating_bins.value_counts().sort_index()

        colors_pie = ['#ff4444', '#ff9944', '#ffdd44', '#88dd44', '#44dd44']
        explode = [0.05 if i == rating_grouped.values.argmax() else 0 for i in range(len(rating_grouped))]

        wedges, texts, autotexts = ax2.pie(rating_grouped.values,
                                           labels=rating_grouped.index,
                                           autopct='%1.1f%%',
                                           colors=colors_pie,
                                           explode=explode,
                                           startangle=90,
                                           textprops={'fontsize': 10, 'fontname': 'Microsoft YaHei'})

        ax2.set_title('星级占比分布', fontsize=14, fontname='Microsoft YaHei')

        for text in texts:
            text.set_fontname('Microsoft YaHei')
        for autotext in autotexts:
            autotext.set_fontname('Microsoft YaHei')
            autotext.set_color('white')
            autotext.set_fontweight('bold')

        plt.tight_layout()
        plt.savefig('rating_distribution_improved.png', dpi=300, bbox_inches='tight')
        print("✓ 已生成: rating_distribution_improved.png")
        plt.close()

    def plot_product_ranking_with_insights(self):
        """产品排名图(带优缺点信息)"""
        fig = plt.figure(figsize=(18, 10))
        gs = fig.add_gridspec(2, 2, hspace=0.3, wspace=0.3)

        fig.suptitle('产品表现排名分析', fontsize=20, fontweight='bold', fontname='Microsoft YaHei')

        product_col = 'Product Category'
        rating_col = 'Rating'

        # 准备数据
        product_stats = self.df.groupby(product_col).agg({
            rating_col: ['mean', 'count']
        }).reset_index()
        product_stats.columns = ['Product', 'AvgRating', 'Count']
        product_stats = product_stats.sort_values('AvgRating', ascending=False)

        # 1. 产品平均评分排名(左上)
        ax1 = fig.add_subplot(gs[0, 0])
        colors = plt.cm.RdYlGn(np.linspace(0.3, 0.9, len(product_stats)))

        bars = ax1.barh(product_stats['Product'], product_stats['AvgRating'], color=colors, edgecolor='black')
        ax1.set_xlabel('平均评分(星)', fontsize=11, fontname='Microsoft YaHei')
        ax1.set_title('产品平均评分排名', fontsize=13, fontweight='bold', fontname='Microsoft YaHei')
        ax1.set_xlim(0, 5.5)
        ax1.grid(axis='x', alpha=0.3)

        for i, (idx, row) in enumerate(product_stats.iterrows()):
            ax1.text(row['AvgRating'] + 0.05, i, f"{row['AvgRating']:.2f}★",
                    va='center', fontsize=9, fontname='Microsoft YaHei')

        # 2. 产品评论数量(右上)
        ax2 = fig.add_subplot(gs[0, 1])
        product_counts = product_stats.sort_values('Count', ascending=False)
        colors2 = plt.cm.Blues(np.linspace(0.4, 0.9, len(product_counts)))

        bars2 = ax2.barh(product_counts['Product'], product_counts['Count'], color=colors2, edgecolor='black')
        ax2.set_xlabel('评论数量', fontsize=11, fontname='Microsoft YaHei')
        ax2.set_title('产品评论数量排名', fontsize=13, fontweight='bold', fontname='Microsoft YaHei')
        ax2.grid(axis='x', alpha=0.3)

        for i, (idx, row) in enumerate(product_counts.iterrows()):
            ax2.text(row['Count'] + 1, i, f"{int(row['Count'])}",
                    va='center', fontsize=9, fontname='Microsoft YaHei')

        # 3. 产品好评率对比(左下)
        ax3 = fig.add_subplot(gs[1, 0])
        good_rates = []
        products_list = []
        for product in product_stats['Product']:
            if product in self.product_insights:
                good_rates.append(self.product_insights[product]['good_review_rate'])
                products_list.append(product)

        colors3 = plt.cm.Greens(np.linspace(0.4, 0.9, len(good_rates)))
        ax3.barh(products_list, good_rates, color=colors3, edgecolor='black')
        ax3.set_xlabel('好评率 (%)', fontsize=11, fontname='Microsoft YaHei')
        ax3.set_title('产品好评率对比 (4星及以上)', fontsize=13, fontweight='bold', fontname='Microsoft YaHei')
        ax3.set_xlim(0, 100)
        ax3.grid(axis='x', alpha=0.3)

        for i, (product, rate) in enumerate(zip(products_list, good_rates)):
            ax3.text(rate + 1, i, f"{rate:.1f}%", va='center', fontsize=9, fontname='Microsoft YaHei')

        # 4. 综合说明(右下)
        ax4 = fig.add_subplot(gs[1, 1])
        ax4.axis('off')
        ax4.text(0.5, 0.95, '📊 产品表现综合评价', ha='center', fontsize=14,
                fontweight='bold', transform=ax4.transAxes, fontname='Microsoft YaHei')

        summary_text = "\n"
        for i, (idx, row) in enumerate(product_stats.iterrows(), 1):
            product = row['Product']
            if product in self.product_insights:
                insight = self.product_insights[product]

                summary_text += f"【{i}. {product}】\n"
                summary_text += f"   评分: {insight['avg_rating']:.2f}★ | 评论: {insight['review_count']}条 | 好评率: {insight['good_review_rate']:.1f}%\n"

                if insight['pros']:
                    top_pro = insight['pros'][0][0]
                    summary_text += f"   ✅ 最大优势: {top_pro}\n"

                if insight['cons']:
                    top_con = insight['cons'][0][0]
                    summary_text += f"   ⚠️  主要问题: {top_con}\n"
                else:
                    summary_text += f"   ✅ 暂无明显问题\n"

                summary_text += "\n"

        ax4.text(0.05, 0.85, summary_text, fontsize=9, verticalalignment='top',
                transform=ax4.transAxes, fontname='Microsoft YaHei',
                bbox=dict(boxstyle='round,pad=1', facecolor='lightblue', alpha=0.2))

        plt.savefig('product_ranking_summary.png', dpi=300, bbox_inches='tight')
        print("✓ 已生成: product_ranking_summary.png")
        plt.close()

    def plot_all_products_details(self):
        """详细的所有产品优缺点对比图"""
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

        fig.suptitle('📋 各产品优缺点详细分析', fontsize=22, fontweight='bold',
                    fontname='Microsoft YaHei', y=0.995)

        for idx, (_, row) in enumerate(product_stats.iterrows()):
            product = row['Product']

            if product not in self.product_insights:
                continue

            insight = self.product_insights[product]

            # 为每个产品创建一个子区域
            ax = plt.subplot(num_products, 1, idx + 1)
            ax.axis('off')

            # 背景色区分
            bg_color = plt.cm.Pastel1(idx % 9)
            rect = plt.Rectangle((0, 0), 1, 1, transform=ax.transAxes,
                                facecolor=bg_color, alpha=0.3, zorder=0)
            ax.add_patch(rect)

            # 产品标题
            title_y = 0.85
            ax.text(0.5, title_y, f"【{product}】", ha='center', fontsize=16,
                   fontweight='bold', transform=ax.transAxes, fontname='Microsoft YaHei',
                   bbox=dict(boxstyle='round,pad=0.8', facecolor='white',
                           edgecolor='black', linewidth=2))

            # 基础数据
            info_y = 0.7
            info_text = f"⭐ 平均评分: {insight['avg_rating']:.2f}★  |  📝 评论数: {insight['review_count']}条  |  👍 好评率: {insight['good_review_rate']:.1f}%"
            ax.text(0.5, info_y, info_text, ha='center', fontsize=12,
                   transform=ax.transAxes, fontname='Microsoft YaHei',
                   bbox=dict(boxstyle='round,pad=0.5', facecolor='lightyellow', alpha=0.8))

            # 优点部分 (左侧)
            pros_y = 0.55
            ax.text(0.05, pros_y, "✅ 主要优点", fontsize=13, fontweight='bold',
                   transform=ax.transAxes, color='darkgreen', fontname='Microsoft YaHei')

            pros_content_y = 0.45
            if insight['pros']:
                pros_text = ""
                for i, (desc, count, kw) in enumerate(insight['pros'][:5], 1):
                    pros_text += f"{i}. {desc} (提及{count}次)\n"
                ax.text(0.08, pros_content_y, pros_text, fontsize=11,
                       transform=ax.transAxes, fontname='Microsoft YaHei',
                       verticalalignment='top',
                       bbox=dict(boxstyle='round,pad=0.8', facecolor='lightgreen', alpha=0.3))
            else:
                ax.text(0.08, pros_content_y, "暂无明显优点关键词", fontsize=10,
                       transform=ax.transAxes, fontname='Microsoft YaHei', style='italic')

            # 缺点部分 (右侧)
            cons_y = 0.55
            ax.text(0.55, cons_y, "⚠️ 主要缺点", fontsize=13, fontweight='bold',
                   transform=ax.transAxes, color='darkred', fontname='Microsoft YaHei')

            cons_content_y = 0.45
            if insight['cons']:
                cons_text = ""
                for i, (desc, count, kw) in enumerate(insight['cons'][:5], 1):
                    cons_text += f"{i}. {desc} (提及{count}次)\n"
                ax.text(0.58, cons_content_y, cons_text, fontsize=11,
                       transform=ax.transAxes, fontname='Microsoft YaHei',
                       verticalalignment='top',
                       bbox=dict(boxstyle='round,pad=0.8', facecolor='#ffcccc', alpha=0.3))
            else:
                ax.text(0.58, cons_content_y, "✨ 暂无明显问题，表现优秀！", fontsize=10,
                       transform=ax.transAxes, fontname='Microsoft YaHei',
                       style='italic', color='green', fontweight='bold')

            # 添加分隔线
            if idx < num_products - 1:
                ax.plot([0, 1], [0.05, 0.05], 'k-', linewidth=2, transform=ax.transAxes)

        plt.tight_layout()
        plt.savefig('all_products_details.png', dpi=300, bbox_inches='tight')
        print("✓ 已生成: all_products_details.png (所有产品详细优缺点)")
        plt.close()

    def run_analysis(self):
        """运行完整分析"""
        if not self.load_data():
            return

        self.analyze_product_pros_cons()
        self.plot_improved_charts()

        print("\n" + "=" * 80)
        print("✅ 分析完成！生成的文件:")
        print("   1. rating_distribution_improved.png - 改进版星级分布图(微软雅黑字体)")
        print("   2. product_ranking_summary.png - 产品排名综合分析图")
        print("   3. all_products_details.png - 所有产品详细优缺点对比图")
        print("=" * 80)

if __name__ == "__main__":
    analyzer = FinalReviewAnalyzer(r"e:\Data\VS\AI\PP\PP\reviews_data.xlsx")
    analyzer.run_analysis()
