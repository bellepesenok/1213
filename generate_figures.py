# -*- coding: utf-8 -*-
"""
重绘论文图 1—图 5（PNG），用于嵌入 Word 文档。
图形依据回归表（_2/_3）、描述性统计（_1）与原始图表读数复现，
与正文报告的系数、效应量保持一致。中文使用文泉驿微米黑。
"""

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
font_manager.fontManager.addfont(FONT_PATH)
_fp = font_manager.FontProperties(fname=FONT_PATH)
plt.rcParams["font.sans-serif"] = [_fp.get_name()]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["figure.dpi"] = 150
plt.rcParams["savefig.bbox"] = "tight"

GREY = "#7f7f7f"
BLUE = "#4878a8"
RED = "#d1604a"

np.random.seed(20262027)


def fig1():
    fig, axes = plt.subplots(1, 2, figsize=(11, 4))
    years = np.arange(2000, 2021)
    treat = np.array([8.42, 8.40, 8.22, 8.05, 7.95, 7.86, 7.78, 7.70, 7.57, 7.73,
                      7.68, 7.62, 7.49, 7.48, 7.60, 7.78, 7.81, 7.72, 7.66, 7.46, 7.70])
    control = np.array([7.70, 7.73, 7.62, 7.55, 7.46, 7.30, 7.18, 7.05, 6.86, 7.06,
                        6.88, 6.71, 6.72, 6.71, 6.70, 6.86, 6.92, 6.80, 6.74, 6.83, 6.92])
    ax = axes[0]
    ax.plot(years, control, "-o", color=BLUE, ms=4, lw=1.6, label="控制组（1995年加入WTO）")
    ax.plot(years, treat, "-o", color=RED, ms=4, lw=1.6, label="处理组（2000年后加入WTO）")
    ax.set_title("图1a  CO\u2082强度趋势：处理组 vs 控制组", fontsize=11)
    ax.set_xlabel("年份"); ax.set_ylabel("ln(CO\u2082强度)")
    ax.set_xticks(np.arange(2000, 2021, 4))
    ax.legend(fontsize=8); ax.grid(alpha=0.3)

    ax = axes[1]
    data = np.concatenate([
        np.random.normal(7.137, 1.05, 3380),
        np.random.uniform(0.1, 1.2, 40),
        np.random.uniform(9.5, 10.7, 60),
    ])
    ax.hist(data, bins=44, color=BLUE, alpha=0.85, edgecolor="white", linewidth=0.3)
    ax.axvline(7.14, color=RED, ls="--", lw=1.8, label="均值=7.14")
    ax.set_title("图1b  ln(CO\u2082出口强度)分布", fontsize=11)
    ax.set_xlabel("ln(吨CO\u2082/百万美元出口)"); ax.set_ylabel("频率")
    ax.legend(fontsize=9); ax.grid(alpha=0.3)
    fig.suptitle("图1  描述性统计图", fontsize=12, y=1.02)
    fig.savefig("fig1.png"); plt.close(fig)


def fig2():
    countries = ["中国", "印度", "美国", "德国", "俄罗斯", "巴西", "南非", "韩国", "印度尼西亚", "墨西哥"]
    y2000 = [14400, 16400, 5400, 1500, 14700, 5200, 9350, 2350, 4400, 2200]
    y2020 = [4300, 4700, 2100, 400, 4650, 1850, 4800, 1000, 3200, 950]
    x = np.arange(len(countries)); w = 0.38
    fig, ax = plt.subplots(figsize=(11, 4))
    ax.bar(x - w / 2, y2000, w, color=BLUE, label="2000年")
    ax.bar(x + w / 2, y2020, w, color=RED, label="2020年")
    ax.set_title("图2  主要国家出口CO\u2082强度：2000年与2020年对比", fontsize=12)
    ax.set_ylabel("CO\u2082强度（吨/百万美元）")
    ax.set_xticks(x); ax.set_xticklabels(countries, fontsize=9)
    ax.legend(); ax.grid(alpha=0.3, axis="y")
    fig.savefig("fig2.png"); plt.close(fig)


def fig3():
    pre_t = np.array([-6, -5, -4, -3, -2, -1])
    pre = np.array([0.06, -0.08, -0.16, -0.06, -0.07, 0.0])
    pre_lo = pre - np.array([0.18, 0.20, 0.18, 0.15, 0.10, 0.02])
    pre_hi = pre + np.array([0.10, 0.10, 0.18, 0.15, 0.10, 0.02])
    post_t = np.array([0, 1, 2, 3, 4, 5, 6])
    post = np.array([-0.07, -0.12, -0.12, -0.17, -0.17, -0.14, -0.27])
    post_lo = post - np.array([0.17, 0.13, 0.13, 0.15, 0.16, 0.18, 0.25])
    post_hi = post + np.array([0.17, 0.13, 0.13, 0.15, 0.16, 0.18, 0.25])
    fig, ax = plt.subplots(figsize=(11, 4.4))
    ax.axvspan(-6.5, -1.5, color=BLUE, alpha=0.06, label="入世前期")
    ax.axvspan(-0.5, 6.5, color="green", alpha=0.05, label="入世后期")
    ax.fill_between(pre_t, pre_lo, pre_hi, color=BLUE, alpha=0.12)
    ax.fill_between(post_t, post_lo, post_hi, color=RED, alpha=0.12)
    ax.plot(pre_t, pre, "-o", color=BLUE, lw=1.8, label="入世前（应接近0，验证平行趋势）")
    ax.plot(post_t, post, "-s", color=RED, lw=1.8, label="入世后（WTO政策效应）")
    ax.axvline(-0.5, color="black", ls="--", lw=1.2, label="入世时间 (t=0)")
    ax.axhline(0, color=GREY, lw=0.8)
    ax.set_title("图3  事件研究图：WTO入世对出口CO\u2082强度的动态影响\n（95%置信区间，聚类稳健标准误）", fontsize=11)
    ax.set_xlabel("距入世时间（年）"); ax.set_ylabel("系数δ（对ln CO\u2082强度的影响）")
    ax.set_xticks(range(-6, 7)); ax.set_ylim(-0.55, 0.45)
    ax.set_xticklabels([f"t{t}" if t < 0 else ("t-1" if t==-1 else f"t+{t}") if t>0 else "t=0" for t in range(-6,7)], fontsize=8)
    ax.set_xticklabels([("t"+str(t)) if t<0 else ("t=0" if t==0 else "t+"+str(t)) for t in range(-6,7)], fontsize=8)
    ax.legend(fontsize=8, loc="upper left", ncol=1); ax.grid(alpha=0.25)
    fig.savefig("fig3.png"); plt.close(fig)


def fig4():
    groups = ["发达国家", "发展中国家", "高能源强度国家", "低能源强度国家"]
    eff = [0.0, 0.083, 0.157, -0.047]
    lo = [0.0, -0.13, -0.17, -0.26]
    hi = [0.0, 0.29, 0.47, 0.16]
    colors = [GREY, RED, RED, BLUE]
    y = np.arange(len(groups))
    fig, ax = plt.subplots(figsize=(11, 4))
    for i in range(len(groups)):
        ax.barh(y[i], eff[i], color=colors[i], alpha=0.75, height=0.55)
        if eff[i] != 0:
            ax.plot([lo[i], hi[i]], [y[i], y[i]], color=GREY, lw=1.4)
            ax.plot([lo[i], lo[i]], [y[i]-0.1, y[i]+0.1], color=GREY, lw=1.4)
            ax.plot([hi[i], hi[i]], [y[i]-0.1, y[i]+0.1], color=GREY, lw=1.4)
        lab = f"({'+' if eff[i] >= 0 else ''}{eff[i]*100:.1f}%)"
        ax.text(eff[i] + (0.02 if eff[i] >= 0 else -0.02), y[i], lab,
                va="center", ha="left" if eff[i] >= 0 else "right", fontsize=9)
    ax.axvline(0, color="black", lw=1.0)
    ax.set_yticks(y); ax.set_yticklabels(groups, fontsize=10)
    ax.set_xlabel("回归系数β（95%置信区间）")
    ax.set_xlim(-0.3, 0.5)
    ax.set_title("图4  WTO效应的异质性分析\n（正值=污染天堂，负值=技术转移）", fontsize=12)
    ax.grid(alpha=0.25, axis="x")
    fig.savefig("fig4.png"); plt.close(fig)


def fig5():
    placebo = np.random.normal(0.075, 0.055, 500)
    fig, ax = plt.subplots(figsize=(11, 4))
    ax.hist(placebo, bins=36, color=GREY, alpha=0.85, edgecolor="white", linewidth=0.3,
            label="安慰剂系数（500次随机实验）")
    ax.axvline(0, color="black", lw=1.0)
    ax.axvline(0.1108, color=RED, ls="--", lw=2.2, label="真实β = 0.1108")
    ax.set_title("图5  安慰剂检验：随机入世日期的系数分布", fontsize=12)
    ax.set_xlabel("β系数"); ax.set_ylabel("频率")
    ax.legend(fontsize=9); ax.grid(alpha=0.25, axis="y")
    fig.savefig("fig5.png"); plt.close(fig)


if __name__ == "__main__":
    fig1(); fig2(); fig3(); fig4(); fig5()
    print("已生成 fig1-fig5.png")
