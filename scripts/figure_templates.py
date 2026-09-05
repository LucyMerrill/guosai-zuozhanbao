# -*- coding: utf-8 -*-
"""figure_templates.py —— 获奖卷图式模板（28 号图式库的可执行版，v1 2026-08-31）

六张卡的可套用函数：B1 栅格矩阵 / C1 正态检验 / C2 MC 叠加 / C3 预测分段 / D1 灵敏度棒图 / A2 柱顶标注
规范：输出宽 14cm（正文宽）、中文 SimHei、图内无标题（图题在 Word 里加）、保存 300dpi PNG。
用法：from figure_templates import grid_matrix, norm_check, mc_overlay, forecast_split, sensitivity_stem, bar_labeled
"""
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

plt.rcParams.update({
    "font.sans-serif": ["SimHei", "Microsoft YaHei"],
    "axes.unicode_minus": False,
    "font.size": 10.5,
})
CM = 1 / 2.54
OUT = "figures_png"
os.makedirs(OUT, exist_ok=True)
PALETTE = {"primary": "#1A6FC4", "second": "#E8743B", "neutral": "#767676",
           "ok": "#3D9970", "warn": "#B22222"}


def _save(fig, name):
    path = os.path.join(OUT, name + ".png")
    fig.savefig(path, dpi=300, bbox_inches="tight")
    plt.close(fig)
    return path


def grid_matrix(matrix, plot_types, crop_names, name="B1_栅格矩阵",
                season="第一季"):
    """B1 地块×作物栅格矩阵（扬大 图2/3 式）。
    matrix: 0/1 数组，形状 (n地块, n作物)；plot_types: 每地块类型名（用于竖线分组）；
    crop_names: 列标签。"""
    mat = np.asarray(matrix)
    n_p, n_c = mat.shape
    fig, ax = plt.subplots(figsize=(14 * CM, max(4, n_p * 0.22 * CM)))
    ax.imshow(mat, cmap="Blues", vmin=0, vmax=1, aspect="auto")
    ax.set_xticks(range(n_c))
    ax.set_xticklabels(crop_names, rotation=90, fontsize=7)
    ax.set_yticks(range(n_p))
    ax.set_yticklabels([f"地块{i+1}" for i in range(n_p)], fontsize=7)
    ax.set_xticks(np.arange(-0.5, n_c, 1), minor=True)
    ax.set_yticks(np.arange(-0.5, n_p, 1), minor=True)
    ax.grid(which="minor", color="#cccccc", linewidth=0.4)
    # 地块类型分界红竖线
    for i in range(1, n_p):
        if plot_types[i] != plot_types[i - 1]:
            ax.axvline(-0.5, color="#cc3333", linewidth=1.2)
    for i in range(1, n_p):
        if plot_types[i] != plot_types[i - 1]:
            ax.axhline(i - 0.5, color="#cc3333", linewidth=0.8)
    ax.set_title(f"{season}种植布局（色块=种植）", fontsize=10)
    return _save(fig, name)


def norm_check(values_list, labels_list, name="C1_正态检验"):
    """C1 正态检验：直方图+拟合曲线+p 值入图例（扬大 图7 式）。
    values_list: [情形1样本, 情形2样本]；labels_list: ["情形(1)", "情形(2)"]"""
    from scipy import stats
    fig, axes = plt.subplots(1, len(values_list),
                             figsize=(7 * CM * len(values_list), 6 * CM))
    axes = np.atleast_1d(axes)
    for ax, vals, lab in zip(axes, values_list, labels_list):
        ax.hist(vals, bins=30, density=True, alpha=0.75,
                color="#5B9BD5", edgecolor="white")
        mu, sd = np.mean(vals), np.std(vals)
        xs = np.linspace(min(vals), max(vals), 200)
        ax.plot(xs, stats.norm.pdf(xs, mu, sd), color="#cc3333", linewidth=1.5)
        _, p = stats.kstest((vals - mu) / sd, "norm")
        ax.legend([f"{lab}（p={p:.3f}）"], loc="upper left", fontsize=9)
        ax.set_xlabel("七年总利润（元）")
        ax.set_ylabel("密度")
    return _save(fig, name)


def mc_overlay(series_list, labels, xlabel="模拟样本号",
               ylabel="七年总利润（元）", name="C2_MC叠加"):
    """C2 蒙特卡洛双方案叠加散点（扬大 图8 式）。
    series_list: [方案A序列, 方案B序列]；labels: 图例名。"""
    fig, ax = plt.subplots(figsize=(14 * CM, 7 * CM))
    colors = ["#1A6FC4", "#E8743B"]
    for i, (s, lab) in enumerate(zip(series_list, labels)):
        ax.plot(range(1, len(s) + 1), s, ".", markersize=2.5,
                color=colors[i % 2], alpha=0.65, label=lab)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.legend(fontsize=9, loc="upper right")
    return _save(fig, name)


def forecast_split(history, forecast, split_idx, ylabel="日销量（千克）",
                   name="C3_预测分段"):
    """C3 预测分段线（C050 图9 式）：历史红实线+预测蓝线+分隔竖线。
    history: 历史序列；forecast: 预测序列；split_idx: 分界位置。"""
    fig, ax = plt.subplots(figsize=(14 * CM, 6 * CM))
    ax.plot(history, color="#cc3333", linewidth=1.0, label="历史")
    ax.plot(range(split_idx - 1, split_idx + len(forecast) - 1), forecast,
            color="#1A6FC4", linewidth=1.4, label="预测")
    ax.axvline(split_idx - 1, color="#222222", linewidth=1.0)
    ax.text(split_idx + 1, ax.get_ylim()[1] * 0.95, "预测", fontsize=9)
    ax.set_ylabel(ylabel)
    ax.legend(fontsize=9)
    return _save(fig, name)


def sensitivity_stem(sens_list, labels, name="D1_灵敏度棒图"):
    """D1 标准化灵敏度 stem 棒图（扬大 图9 式），两情景并排，0 参考线。
    sens_list: [情景A灵敏度序列, 情景B灵敏度序列]。"""
    fig, axes = plt.subplots(1, len(sens_list),
                             figsize=(7 * CM * len(sens_list), 6 * CM))
    axes = np.atleast_1d(axes)
    for ax, s, lab in zip(axes, sens_list, labels):
        ax.stem(range(1, len(s) + 1), s, basefmt="k-")
        ax.axhline(0, color="#222222", linewidth=0.8)
        ax.set_xlabel("样本编号")
        ax.set_ylabel("标准化灵敏度")
        ax.set_title(lab, fontsize=10)
    return _save(fig, name)


def bar_labeled(categories, values, ylabel="数值", name="A2_柱顶标注"):
    """A2 柱状图+柱顶数值标注（C050 图4 式）。"""
    fig, ax = plt.subplots(figsize=(14 * CM, 7 * CM))
    bars = ax.bar(categories, values, color="#5B9BD5", width=0.55)
    for b, v in zip(bars, values):
        ax.text(b.get_x() + b.get_width() / 2, v, f"{v:g}",
                ha="center", va="bottom", fontsize=9)
    ax.set_ylabel(ylabel)
    return _save(fig, name)


if __name__ == "__main__":
    # 冒烟：随机数据各渲染一张，验证管线
    rng = np.random.default_rng(1)
    grid_matrix(rng.integers(0, 2, (26, 15)),
                ["平旱地"] * 6 + ["梯田"] * 14 + ["山坡地"] * 6,
                [f"作物{i+1}" for i in range(15)])
    norm_check([rng.normal(4.1e7, 1e5, 1000), rng.normal(4.11e7, 8e4, 1000)],
               ["情形(1)", "情形(2)"])
    mc_overlay([rng.normal(6.22e7, 7e5, 1000), rng.normal(6.23e7, 7e5, 1000)],
               ["方案一", "方案二"])
    forecast_split(rng.gamma(2, 60, 300).tolist(), [200, 210, 190], 299)
    sensitivity_stem([rng.normal(0, 1, 100), rng.normal(0, 1, 100)],
                     ["情景 1", "情景 2"])
    bar_labeled([f"品类{i}" for i in range(6)], rng.integers(5, 15, 6))
    print("模板冒烟通过，输出在", OUT)
