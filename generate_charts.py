"""
ネクプロ戦略プレゼンテーション用チャート画像生成
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
import os

# --- Font setup ---
JP_FONT = None
for candidate in ['WenQuanYi Zen Hei', 'Noto Sans CJK JP', 'IPAGothic', 'unifont_jp']:
    try:
        fp = fm.findfont(fm.FontProperties(family=candidate))
        if fp and 'last' not in fp.lower():
            JP_FONT = candidate
            break
    except Exception:
        continue

if JP_FONT:
    plt.rcParams['font.family'] = JP_FONT
else:
    plt.rcParams['font.family'] = 'sans-serif'

plt.rcParams['axes.unicode_minus'] = False

OUT_DIR = '/home/user/nexpro/chart_images'
os.makedirs(OUT_DIR, exist_ok=True)

# Color palette - clean corporate (navy, blue, grey, accent)
NAVY = '#1B2A4A'
BLUE = '#2E6DA4'
LIGHT_BLUE = '#5BA4CF'
ACCENT = '#E8913A'
ACCENT_RED = '#D64045'
GREY = '#8C8C8C'
LIGHT_GREY = '#D9D9D9'
WHITE = '#FFFFFF'
BG_COLOR = '#FAFAFA'

def save(fig, name):
    fig.savefig(f'{OUT_DIR}/{name}', dpi=200, bbox_inches='tight',
                facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close(fig)
    print(f'  Saved: {name}')


# ========== Chart 1: Revenue Trend ==========
def chart_revenue_trend():
    fig, ax = plt.subplots(figsize=(10, 5.5), facecolor=BG_COLOR)
    ax.set_facecolor(BG_COLOR)

    years = ['FY22', 'FY23', 'FY24', 'FY25\n(計画)', 'FY26\n(計画)', 'FY27\n(計画)']
    revenue = [418, 497, 513, 650, 912, 1383]
    mrr = [226, 253, 288, 330, 417, 527]
    option_svc = [192, 245, 227, 320, 495, 856]

    x = np.arange(len(years))
    w = 0.35

    bars1 = ax.bar(x - w/2, mrr, w, label='MRR（システム利用料）', color=NAVY, zorder=3)
    bars2 = ax.bar(x + w/2, option_svc, w, label='オプション+新規事業', color=LIGHT_BLUE, zorder=3)

    ax.plot(x, revenue, color=ACCENT, marker='o', markersize=8, linewidth=2.5,
            label='売上総合計', zorder=4)

    for i, v in enumerate(revenue):
        ax.annotate(f'¥{v:,}M', (x[i], v), textcoords="offset points",
                    xytext=(0, 12), ha='center', fontsize=10, fontweight='bold',
                    color=ACCENT)

    # YoY labels
    yoy = [None, '+19.0%', '+3.1%', '+26.7%', '+40.3%', '+51.6%']
    for i, y in enumerate(yoy):
        if y:
            color = ACCENT_RED if y == '+3.1%' else NAVY
            ax.annotate(y, (x[i], revenue[i]), textcoords="offset points",
                        xytext=(0, 26), ha='center', fontsize=8, color=color)

    ax.set_xticks(x)
    ax.set_xticklabels(years, fontsize=10)
    ax.set_ylabel('百万円 (M)', fontsize=10, color=GREY)
    ax.set_title('売上推移と構成（FY22-FY27）', fontsize=14, fontweight='bold',
                 color=NAVY, pad=20)
    ax.legend(loc='upper left', fontsize=9, framealpha=0.9)
    ax.grid(axis='y', alpha=0.3, zorder=0)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.set_ylim(0, 1600)

    save(fig, 'revenue_trend.png')


# ========== Chart 2: MRR + ARPA Dual Axis ==========
def chart_mrr_arpa():
    fig, ax1 = plt.subplots(figsize=(10, 5), facecolor=BG_COLOR)
    ax1.set_facecolor(BG_COLOR)

    years = ['FY22', 'FY23', 'FY24', 'FY25(計画)', 'FY26(計画)', 'FY27(計画)']
    mrr_annual = [226, 253, 288, 330, 417, 527]
    arpa = [105, 138, 148, 169, 186, 204]

    x = np.arange(len(years))

    bars = ax1.bar(x, mrr_annual, 0.5, color=NAVY, alpha=0.85, label='MRR年間合計(M)', zorder=3)
    ax1.set_ylabel('MRR年間合計（百万円）', color=NAVY, fontsize=10)
    ax1.set_ylim(0, 650)

    ax2 = ax1.twinx()
    ax2.plot(x, arpa, color=ACCENT, marker='s', markersize=8, linewidth=2.5,
             label='ARPA長期PF(千円/月)', zorder=4)
    ax2.set_ylabel('ARPA（千円/月）', color=ACCENT, fontsize=10)
    ax2.set_ylim(50, 250)

    for i, v in enumerate(arpa):
        ax2.annotate(f'¥{v}K', (x[i], v), textcoords="offset points",
                     xytext=(0, 10), ha='center', fontsize=9, fontweight='bold',
                     color=ACCENT)

    ax1.set_xticks(x)
    ax1.set_xticklabels(years, fontsize=9)
    ax1.set_title('MRR成長とARPA推移', fontsize=14, fontweight='bold', color=NAVY, pad=15)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left', fontsize=9)

    ax1.grid(axis='y', alpha=0.3, zorder=0)
    ax1.spines['top'].set_visible(False)

    save(fig, 'mrr_arpa.png')


# ========== Chart 3: New Revenue Streams ==========
def chart_new_revenue():
    fig, ax = plt.subplots(figsize=(10, 5), facecolor=BG_COLOR)
    ax.set_facecolor(BG_COLOR)

    years = ['FY25(計画)', 'FY26(計画)', 'FY27(計画)']
    compound = [3.2, 17, 51]
    sales_dx = [33.5, 121.8, 327]
    x = np.arange(len(years))
    w = 0.35

    ax.bar(x - w/2, compound, w, label='コンパウンド\n(AF・名刺OCR・動画マニュアル)', color=LIGHT_BLUE, zorder=3)
    ax.bar(x + w/2, sales_dx, w, label='営業DX\n(営業代行・営業DX)', color=ACCENT, zorder=3)

    for i in range(len(years)):
        ax.annotate(f'¥{compound[i]}M', (x[i] - w/2, compound[i]),
                    textcoords="offset points", xytext=(0, 5), ha='center', fontsize=9, color=NAVY)
        ax.annotate(f'¥{sales_dx[i]}M', (x[i] + w/2, sales_dx[i]),
                    textcoords="offset points", xytext=(0, 5), ha='center', fontsize=9,
                    fontweight='bold', color=ACCENT)

    # Total labels
    for i in range(len(years)):
        total = compound[i] + sales_dx[i]
        pct_of_total = [5.6, 15.2, 27.3][i]
        ax.annotate(f'合計 ¥{total:.0f}M\n(全体の{pct_of_total}%)',
                    (x[i], max(compound[i], sales_dx[i])),
                    textcoords="offset points", xytext=(0, 20), ha='center',
                    fontsize=9, fontweight='bold', color=NAVY,
                    bbox=dict(boxstyle='round,pad=0.3', facecolor='#E8F4FD', edgecolor=LIGHT_BLUE, alpha=0.8))

    ax.set_xticks(x)
    ax.set_xticklabels(years, fontsize=11)
    ax.set_ylabel('百万円 (M)', fontsize=10, color=GREY)
    ax.set_title('新収益柱の成長計画', fontsize=14, fontweight='bold', color=NAVY, pad=20)
    ax.legend(loc='upper left', fontsize=9)
    ax.grid(axis='y', alpha=0.3, zorder=0)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.set_ylim(0, 420)

    save(fig, 'new_revenue.png')


# ========== Chart 4: Churn Rate Trend ==========
def chart_churn():
    fig, ax = plt.subplots(figsize=(10, 4.5), facecolor=BG_COLOR)
    ax.set_facecolor(BG_COLOR)

    years = ['FY22', 'FY23', 'FY24', 'FY25(目標)', 'FY26(目標)', 'FY27(目標)']
    churn = [3.6, 2.3, 1.7, 1.0, 1.0, 1.0]
    colors = [ACCENT_RED, ACCENT, ACCENT, BLUE, BLUE, BLUE]

    bars = ax.bar(years, churn, 0.5, color=colors, zorder=3, alpha=0.85)

    for bar, v in zip(bars, churn):
        ax.annotate(f'{v}%', (bar.get_x() + bar.get_width()/2, v),
                    textcoords="offset points", xytext=(0, 8), ha='center',
                    fontsize=12, fontweight='bold', color=NAVY)

    ax.axhline(y=1.0, color=BLUE, linestyle='--', alpha=0.5, linewidth=1.5, label='目標: 1.0%')
    ax.axhline(y=0.42, color='green', linestyle=':', alpha=0.5, linewidth=1.5, label='SaaS優良水準: ~0.4%/月(年5%)')

    ax.set_title('月次解約率（長期PF）推移と目標', fontsize=14, fontweight='bold', color=NAVY, pad=15)
    ax.set_ylabel('月次解約率 (%)', fontsize=10, color=GREY)
    ax.legend(fontsize=9, loc='upper right')
    ax.grid(axis='y', alpha=0.3, zorder=0)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.set_ylim(0, 5)

    save(fig, 'churn_rate.png')


# ========== Chart 5: Positioning Map 1 ==========
def chart_positioning_map1():
    fig, ax = plt.subplots(figsize=(9, 7), facecolor=BG_COLOR)
    ax.set_facecolor(BG_COLOR)

    companies = {
        'Zoom Webinars': (7.5, 3.0, GREY),
        'ON24': (8.5, 3.5, GREY),
        'EventHub': (4.5, 7.5, LIGHT_BLUE),
        'bizibl': (3.5, 7.0, LIGHT_BLUE),
        'FanGrowth': (3.0, 8.0, LIGHT_BLUE),
        'ネクプロ\n(現在)': (6.0, 7.5, BLUE),
        'ネクプロ\n(目標)': (8.5, 9.0, ACCENT),
    }

    for name, (x, y, color) in companies.items():
        size = 220 if 'ネクプロ' in name else 150
        marker = '★' if 'ネクプロ' in name else 'o'
        if 'ネクプロ' in name:
            ax.scatter(x, y, s=size, c=color, zorder=5, marker='*' if '目標' in name else 'o',
                       edgecolors='black', linewidths=0.5)
        else:
            ax.scatter(x, y, s=size, c=color, zorder=5, edgecolors='black', linewidths=0.5)
        offset_y = 15 if '目標' not in name else -20
        ax.annotate(name, (x, y), textcoords="offset points", xytext=(12, offset_y),
                    fontsize=10, fontweight='bold' if 'ネクプロ' in name else 'normal',
                    color=color)

    # Arrow from current to target
    ax.annotate('', xy=(8.3, 8.8), xytext=(6.2, 7.7),
                arrowprops=dict(arrowstyle='->', color=ACCENT, lw=2.5, linestyle='--'))

    ax.set_xlabel('機能深度（配信 + 分析 + 実行）→', fontsize=11, color=NAVY, fontweight='bold')
    ax.set_ylabel('日本企業適合性（商習慣 / 支援 / 言語 / 連携）→', fontsize=11, color=NAVY, fontweight='bold')
    ax.set_title('Positioning Map 1: 機能深度 × 日本企業適合性',
                 fontsize=13, fontweight='bold', color=NAVY, pad=15)

    ax.set_xlim(1, 10.5)
    ax.set_ylim(1, 10.5)
    ax.axvline(x=5.5, color=LIGHT_GREY, linestyle='-', alpha=0.5)
    ax.axhline(y=5.5, color=LIGHT_GREY, linestyle='-', alpha=0.5)
    ax.grid(alpha=0.15, zorder=0)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    # Quadrant labels
    ax.text(3, 9.8, '国内ツール型\n(高適合・低機能)', fontsize=8, color=GREY, ha='center', style='italic')
    ax.text(9, 9.8, '目標ポジション\n(高適合・高機能)', fontsize=8, color=ACCENT, ha='center', fontweight='bold')
    ax.text(3, 1.5, '汎用ツール\n(低適合・低機能)', fontsize=8, color=GREY, ha='center', style='italic')
    ax.text(9, 1.5, 'グローバル専業\n(低適合・高機能)', fontsize=8, color=GREY, ha='center', style='italic')

    save(fig, 'positioning_map1.png')


# ========== Chart 6: Positioning Map 2 ==========
def chart_positioning_map2():
    fig, ax = plt.subplots(figsize=(9, 7), facecolor=BG_COLOR)
    ax.set_facecolor(BG_COLOR)

    companies = {
        'Zoom Webinars': (3.0, 3.5, GREY),
        'ON24': (8.5, 7.5, GREY),
        'EventHub': (3.5, 4.5, LIGHT_BLUE),
        'bizibl': (2.5, 3.0, LIGHT_BLUE),
        'FanGrowth': (2.0, 2.5, LIGHT_BLUE),
        'ネクプロ\n(現在)': (5.0, 5.0, BLUE),
        'ネクプロ\n(目標)': (8.0, 3.5, ACCENT),
    }

    for name, (x, y, color) in companies.items():
        size = 220 if 'ネクプロ' in name else 150
        if 'ネクプロ' in name:
            ax.scatter(x, y, s=size, c=color, zorder=5, marker='*' if '目標' in name else 'o',
                       edgecolors='black', linewidths=0.5)
        else:
            ax.scatter(x, y, s=size, c=color, zorder=5, edgecolors='black', linewidths=0.5)
        offset_y = 15 if '目標' not in name else -20
        ax.annotate(name, (x, y), textcoords="offset points", xytext=(12, offset_y),
                    fontsize=10, fontweight='bold' if 'ネクプロ' in name else 'normal',
                    color=color)

    ax.annotate('', xy=(7.8, 3.7), xytext=(5.2, 4.8),
                arrowprops=dict(arrowstyle='->', color=ACCENT, lw=2.5, linestyle='--'))

    # Sweet spot highlight
    from matplotlib.patches import Ellipse
    ellipse = Ellipse((8.0, 3.5), 3.0, 2.5, alpha=0.08, color=ACCENT, zorder=1)
    ax.add_patch(ellipse)
    ax.text(8.0, 2.0, 'Sweet Spot', fontsize=9, color=ACCENT, ha='center', fontweight='bold')

    ax.set_xlabel('データ活用高度性（記録 → 示唆 → 自動実行）→', fontsize=11, color=NAVY, fontweight='bold')
    ax.set_ylabel('← 導入ハードル（低い方が良い）', fontsize=11, color=NAVY, fontweight='bold')
    ax.set_title('Positioning Map 2: データ活用高度性 × 導入ハードル',
                 fontsize=13, fontweight='bold', color=NAVY, pad=15)

    ax.set_xlim(0.5, 10.5)
    ax.set_ylim(0.5, 10.5)
    ax.invert_yaxis()
    ax.axvline(x=5.5, color=LIGHT_GREY, linestyle='-', alpha=0.5)
    ax.axhline(y=5.5, color=LIGHT_GREY, linestyle='-', alpha=0.5)
    ax.grid(alpha=0.15, zorder=0)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    save(fig, 'positioning_map2.png')


# ========== Chart 7: SaaS 3-Layer Structure ==========
def chart_saas_layers():
    fig, ax = plt.subplots(figsize=(10, 5.5), facecolor=BG_COLOR)
    ax.set_facecolor(BG_COLOR)

    layers = [
        ('AIエージェント層（実行）', 3.0, BLUE, WHITE, '価値増大 — 業務を自律実行'),
        ('SaaS UIレイヤー（中間層）', 2.0, ACCENT_RED, WHITE, '圧縮対象 — ダッシュボード・ワークフロー'),
        ('System of Record（データ層）', 1.0, NAVY, WHITE, '価値増大 — CRM・ERP・独自データ'),
    ]

    for label, y, color, text_color, desc in layers:
        width = 6
        rect = plt.Rectangle((2, y - 0.35), width, 0.7, facecolor=color, edgecolor='white',
                              linewidth=2, zorder=3, alpha=0.9)
        ax.add_patch(rect)
        ax.text(5, y + 0.05, label, ha='center', va='center', fontsize=12,
                fontweight='bold', color=text_color, zorder=4)
        ax.text(5, y - 0.2, desc, ha='center', va='center', fontsize=8,
                color=text_color, alpha=0.85, zorder=4)

    # Squeeze arrows
    ax.annotate('', xy=(8.5, 2.35), xytext=(8.5, 2.7),
                arrowprops=dict(arrowstyle='->', color=ACCENT_RED, lw=3))
    ax.annotate('', xy=(8.5, 1.65), xytext=(8.5, 1.3),
                arrowprops=dict(arrowstyle='->', color=ACCENT_RED, lw=3))
    ax.text(9.2, 2.0, '圧縮', fontsize=11, color=ACCENT_RED, fontweight='bold',
            ha='center', va='center')

    # Nexpro position
    ax.annotate('ネクプロの\n目指す位置', xy=(1.8, 1.0), xytext=(0.3, 0.3),
                fontsize=10, fontweight='bold', color=ACCENT,
                arrowprops=dict(arrowstyle='->', color=ACCENT, lw=2),
                bbox=dict(boxstyle='round,pad=0.3', facecolor='#FFF3E0', edgecolor=ACCENT))

    ax.set_xlim(-0.5, 10.5)
    ax.set_ylim(0, 4.2)
    ax.set_title('AIエージェント時代のSaaS価値構造', fontsize=14, fontweight='bold', color=NAVY, pad=10)
    ax.axis('off')

    save(fig, 'saas_layers.png')


# ========== Chart 8: Roadmap Timeline ==========
def chart_roadmap():
    fig, ax = plt.subplots(figsize=(12, 6), facecolor=BG_COLOR)
    ax.set_facecolor(BG_COLOR)

    phases = [
        ('短期 0-6M\n止血・改善', 0, 6, ACCENT_RED),
        ('中期 6-18M\n転換・仕込み', 6, 12, ACCENT),
        ('長期 18-36M\n成長・回収', 18, 18, BLUE),
    ]

    tracks = {
        'プロダクト': [
            ('エンゲージメントスコアMVP', 0, 6, NAVY),
            ('AI コンテンツ生成', 4, 8, NAVY),
            ('API-first移行', 6, 12, NAVY),
            ('HubSpot/Marketo連携', 18, 8, NAVY),
        ],
        'GTM': [
            ('業種別パッケージ', 0, 4, BLUE),
            ('価格体系再設計', 6, 6, BLUE),
            ('営業DX加速', 0, 18, BLUE),
            ('ブランド・リポジショニング', 18, 12, BLUE),
        ],
        'CS': [
            ('オンボーディング標準化', 0, 3, LIGHT_BLUE),
            ('ヘルススコア導入', 2, 4, LIGHT_BLUE),
            ('CS Profit Center化', 6, 8, LIGHT_BLUE),
            ('戦略アカウント制', 0, 6, LIGHT_BLUE),
        ],
        '組織': [
            ('PMM兼務設置', 0, 3, GREY),
            ('RevOps設置', 3, 6, GREY),
            ('戦略採用 3-5名', 6, 8, GREY),
            ('KPIオーナー制度', 0, 1, GREY),
        ],
    }

    track_names = list(tracks.keys())
    y_positions = {name: i * 2.5 for i, name in enumerate(reversed(track_names))}

    # Phase backgrounds
    phase_y_min = -0.8
    phase_y_max = max(y_positions.values()) + 1.8
    for label, start, duration, color in phases:
        ax.axvspan(start, start + duration, alpha=0.06, color=color, zorder=0)
        ax.text(start + duration/2, phase_y_max + 0.3, label,
                ha='center', va='bottom', fontsize=9, fontweight='bold', color=color)

    # Track items
    for track_name, items in tracks.items():
        y = y_positions[track_name]
        ax.text(-1.5, y + 0.3, track_name, ha='right', va='center',
                fontsize=11, fontweight='bold', color=NAVY)
        for i, (item_label, start, duration, color) in enumerate(items):
            bar_y = y + (i % 2) * 0.6
            ax.barh(bar_y, duration, left=start, height=0.45, color=color, alpha=0.7,
                    edgecolor='white', linewidth=1, zorder=3)
            text_x = start + duration / 2
            ax.text(text_x, bar_y, item_label, ha='center', va='center',
                    fontsize=7, color=WHITE, fontweight='bold', zorder=4)

    # Gate Review markers
    for month, label in [(6, 'Gate\nReview'), (12, 'Gate\nReview'), (18, 'Gate\nReview')]:
        ax.axvline(x=month, color=ACCENT_RED, linestyle='--', alpha=0.4, zorder=1)
        ax.text(month, phase_y_max - 0.3, label, ha='center', fontsize=7,
                color=ACCENT_RED, fontweight='bold')

    ax.set_xlim(-2, 37)
    ax.set_ylim(phase_y_min, phase_y_max + 1.5)
    ax.set_xlabel('月数', fontsize=10, color=GREY)
    ax.set_title('3層実行ロードマップ（0-36ヶ月）', fontsize=14, fontweight='bold', color=NAVY, pad=20)
    ax.set_xticks([0, 6, 12, 18, 24, 30, 36])
    ax.set_xticklabels(['0', '6', '12', '18', '24', '30', '36'])
    ax.set_yticks([])
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)

    save(fig, 'roadmap.png')


# ========== Chart 9: KPI Tree ==========
def chart_kpi_tree():
    fig, ax = plt.subplots(figsize=(11, 6.5), facecolor=BG_COLOR)
    ax.set_facecolor(BG_COLOR)

    def draw_box(x, y, text, color, w=2.2, h=0.55, fontsize=8):
        rect = plt.Rectangle((x - w/2, y - h/2), w, h, facecolor=color,
                              edgecolor='white', linewidth=1.5, zorder=3, alpha=0.9)
        ax.add_patch(rect)
        ax.text(x, y, text, ha='center', va='center', fontsize=fontsize,
                fontweight='bold', color=WHITE, zorder=4)

    def draw_line(x1, y1, x2, y2):
        ax.plot([x1, x2], [y1, y2], color=LIGHT_GREY, linewidth=1.5, zorder=2)

    # North Star
    draw_box(5, 5.5, 'North Star\n顧客あたりエンゲージメント成果価値', ACCENT, w=4, h=0.7, fontsize=9)

    # Level 2 - Business KPIs
    biz_kpis = [
        (1.5, 4.0, 'ARR成長率\n+27%→+40%→+52%'),
        (4.2, 4.0, 'NRR\n100%→110%(仮説)'),
        (6.8, 4.0, '粗利率'),
        (9.5, 4.0, '顧客基盤\n167→210社'),
    ]
    for x, y, text in biz_kpis:
        draw_box(x, y, text, NAVY)
        draw_line(5, 5.15, x, 4.28)

    # Level 3 - Leading KPIs
    leading = [
        (0.3, 2.5, 'MRR\n¥330M→¥527M'),
        (2.3, 2.5, '新規MRR\n(長期PF)'),
        (4.2, 2.5, 'Expansion\nMRR'),
        (6.0, 2.5, 'Churn MRR\n解約率1.0%'),
        (7.8, 2.5, 'ARPA\n¥169K→¥204K'),
        (9.8, 2.5, '成約率\n12%→15%'),
    ]
    connections = [(0, 1.5), (1, 1.5), (2, 4.2), (3, 4.2), (4, 6.8), (5, 9.5)]
    for i, (x, y, text) in enumerate(leading):
        draw_box(x, y, text, BLUE, w=1.7, h=0.55, fontsize=7)
        parent_x = connections[i][1]
        draw_line(parent_x, 3.72, x, 2.78)

    # Level 4 - New Revenue
    new_rev = [
        (1.5, 1.2, '営業DX\n¥33M→¥327M'),
        (4.0, 1.2, 'コンパウンド\n¥3M→¥51M'),
        (7.0, 1.2, 'エンゲージメント\nスコア導入数'),
        (9.5, 1.2, 'オンボード\n完了率90%'),
    ]
    new_connections = [(0, 1.5), (1, 1.5), (2, 9.5), (3, 9.5)]
    for i, (x, y, text) in enumerate(new_rev):
        draw_box(x, y, text, LIGHT_BLUE, w=1.9, h=0.55, fontsize=7)

    ax.set_xlim(-1, 11)
    ax.set_ylim(0.5, 6.3)
    ax.set_title('KPIツリー構造', fontsize=14, fontweight='bold', color=NAVY, pad=10)
    ax.axis('off')

    save(fig, 'kpi_tree.png')


# ========== Chart 10: Account Trend ==========
def chart_accounts():
    fig, ax = plt.subplots(figsize=(10, 5), facecolor=BG_COLOR)
    ax.set_facecolor(BG_COLOR)

    years = ['FY22', 'FY23', 'FY24', 'FY25(計画)', 'FY26(計画)', 'FY27(計画)']
    long_term = [160, 151, 167, 179, 195, 210]
    new_per_year = [60, 27, 50, 38, 38, 38]

    x = np.arange(len(years))

    ax.bar(x, long_term, 0.5, color=NAVY, alpha=0.85, label='累計長期PFアカウント数', zorder=3)

    ax2 = ax.twinx()
    ax2.plot(x, new_per_year, color=ACCENT, marker='D', markersize=7, linewidth=2,
             label='年間新規獲得数', zorder=4)

    for i, v in enumerate(long_term):
        ax.annotate(f'{v}社', (x[i], v), textcoords="offset points",
                    xytext=(0, 8), ha='center', fontsize=10, fontweight='bold', color=NAVY)

    for i, v in enumerate(new_per_year):
        ax2.annotate(f'{v}社', (x[i], v), textcoords="offset points",
                     xytext=(0, 10), ha='center', fontsize=9, color=ACCENT, fontweight='bold')

    ax.set_xticks(x)
    ax.set_xticklabels(years, fontsize=9)
    ax.set_ylabel('累計アカウント数', color=NAVY, fontsize=10)
    ax2.set_ylabel('年間新規獲得数', color=ACCENT, fontsize=10)
    ax.set_ylim(0, 260)
    ax2.set_ylim(0, 80)

    ax.set_title('長期PFアカウント数推移', fontsize=14, fontweight='bold', color=NAVY, pad=15)

    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(lines1 + lines2, labels1 + labels2, loc='upper left', fontsize=9)

    ax.grid(axis='y', alpha=0.3, zorder=0)
    ax.spines['top'].set_visible(False)

    save(fig, 'accounts.png')


# ========== Generate All ==========
if __name__ == '__main__':
    print('Generating charts...')
    chart_revenue_trend()
    chart_mrr_arpa()
    chart_new_revenue()
    chart_churn()
    chart_positioning_map1()
    chart_positioning_map2()
    chart_saas_layers()
    chart_roadmap()
    chart_kpi_tree()
    chart_accounts()
    print(f'All charts saved to {OUT_DIR}/')
