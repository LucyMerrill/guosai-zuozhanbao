---
name: guosai-zuozhanbao
description: 数学建模国赛（CUMCM）/华数杯作战包。参加数学建模竞赛、写建模论文、改进已有建模论文时使用。涵盖破题契约→建模→代码求解→写论文→自动评审循环→Word docx 编译提交的全流程 SOP 与技能指令。
---

# 国赛作战包


- `references/00_作战SOP.md` — **总入口**：6 阶段流水线 + 改进论文模式。执行任何比赛任务前先读它。
- `references/21_赛题识别八类任务判据.md` — **破题/选型圣经**（①②用）：开赛40分钟流程、八类任务判据+必报量、四组易混、外推10%判据、优化规模分档与gap、线性化三招、整卷统筹两接口、六类高频误配、选题概率账。
- `references/22_题型流程化解题手册.md` — **题型作业手册**（③④用）：相关/评价/聚类/分类/回归预测/ARIMA/优化/降维完整流程，模型选择条件表+必做检验速查+不推荐清单。
- `references/23_绘图公共头部与生成流程借鉴.md` — 绘图公共头部代码（中文字体/300dpi/SVG 出口/流程图固定模板）+ 生成流程机制（先算后画、防绕圈、预读缓存、摘要字数分配），已分阶段接入 stage1~4；本文存原理与代码底稿。
- `references/24_数字诚实性与评审循环雷点.md` — **数字与逻辑自洽雷点库**（20 号姊妹篇）：叙事与 JSON 冲突、样本量多口径、表述与实现不符、图表映射错位等实战沉淀 + 评审循环换视角操作法（⑤用）。
- `references/25_流程图绘制SOP.md` — 论文流程图 draw.io 路线（④ 用）：**全卷一张 C050 式总体流程图**（§0B：每格实货/与摘要同构/检验不上图，分问图默认不画）+ 三种版式模板（`assets/示例_数模流程图风格.drawio`）+ CLI 逐页导出脚本（`scripts/export_drawio_png.py`，本机装在 D:\drawio）；比 23 号 §3 的 matplotlib 3×3 模板更接近国奖观感，优先用，matplotlib 兜底。
- `references/26_图表模式库.md` — 数据图按题型抄的可运行模式库（③④ 用）：评价/预测/优化/对比稳健/数据探索 7 类，本地化适配中文标签、无图内标题、组图 ≤2；配套 `23` 号 §2.5 成图验收清单。
- `references/30_数模模型字典.xlsx`（原件）+ `31_模型字典_全审标注版.xlsx` + `30_模型字典使用说明与审核结论.md` — **单模型深查字典，已于 2026-09-03 全量审核 5713 条**（可信 74.9%/注水 18.6%/存疑 5.8%/错误 0.5%/未核验 0.2%）：选型写假设定检验只用「可信」条目，比赛期间无需再自行核验。用 `scripts/query_model_dict.py` 检索（带审核结论，`--only-ok` 过滤）。
- `references/32_命题人评阅要点库.md` — **命题人视角评分逻辑**（①~⑤ 全程用）：2020~2025 官方命题人解析 22 篇+赛区评分细则提炼——14 条跨题铁律（结果合理性量级审查/能枚举不启发式/智能算法三件套/隐藏条件用足/模型符号化/检验是显式得分项等）、分题型命题人偏好、逐题失分点速查表（模拟训练对照）、2025 赛区逐问分值骨架（摘要 10 分固定+每问检验档+创新加分推荐机制）。原件在工作区 C题_作战/官方资料_命题人评阅/。
- `references/官方文件/` — **组委会官方原件**（裁决标准，高于一切二手整理）：《格式规范（2026 修订稿）》《AI 工具使用规定（2026 试行）》、2026 标准论文 Word 模板（WPS 制作，仅供参考）。17 号 §0B 硬性红线即从格式规范提炼。
- `references/01~03` — 研究契约模板、Claims-Evidence 矩阵、实验计划模板（①②阶段产出物）
- `references/04_写作铁律_去AI痕迹.md` — 写论文（④）的 7 条铁律
- `references/05_自动评审清单.md` / `12_123条全流程自查清单_原文.md` — ⑤ 自动评审循环用
- `references/06_改进论文流程.md` — 特殊模式：输入原优秀论文改出改进稿
- `references/07_技能_comp-modeling.md` — ② 建模阶段的执行指令
- `references/08_技能_auto-review.md` — ⑤ 评审阶段的执行指令
- `references/09_技能_comp-code.md` — ③ 代码求解阶段的执行指令
- `references/10_技能_comp-paper.md` — ④ 论文撰写阶段的执行指令
- `references/13_比赛实战说明书_新窗口也能跑通.md` — 比赛当天实操手册
- `references/13b_选题决策打分表.md` — 开赛选题决策
- `references/14~18` — 编译与提交说明、国赛 Word 原模板格式规范、程序化校验说明
- `references/11,16,19,20` — 获奖论文学习报告、华数杯临时配置、2025C题获奖卷对标与雷点库等参考材料

## 配套主力技能（已装在本目录同级，SOP 中说的"已装 skill"即它们）

- `math-modeling-solver` — 方法主力：12 类问题本质 + 95+ 模型决策矩阵 + 5 本算法 Cookbook + 11 本 Playbook + 22 个 Python/7 个 MATLAB 代码模板（②选模型、③写代码时调用）
- `math-modeling-paper` — 论文写作主力：章节结构、摘要写法、模型检验、灵敏度分析（④写论文时调用）
- `math-modeling-review` — 国奖评审打分：逐项对照国奖标准出分项评分表（⑤自动评审循环中叠加使用）
- 赛题 PDF 转 md：用 `markdown-converter` 技能（`uvx markitdown 题目.pdf -o 题目.md`）；扫描件 OCR 见 document-skills 的 pdf 技能

## 执行方式：子 agent 流水线（默认）

单个 agent 从头跑到尾上下文必爆。本 skill 的 `stages/` 目录下已备好每阶段的**现成子agent提示词模板**：

```
stages/stage1_破题与契约.md   → general-purpose 子agent
stages/stage2_建模.md         → general-purpose 子agent（可按问题拆）
stages/stage3_代码求解.md     → 每问一个并行子agent
stages/stage4_写论文.md       → 按 A/B/C 批次分派多个写作子agent
stages/stage5_自动评审循环.md → 评审者与改稿者必须是两个不同子agent，≤3轮
stages/stage6_交付编译.md     → general-purpose 子agent
stages/stage7_赛后复盘.md     → 主 agent 亲自执行（赛后48小时内，战果回流）
```

**主 agent（你自己）只做调度，不亲自干活：**

1. 准备期：把题目 PDF 用 `uvx markitdown 题目.pdf -o 工作目录/题目.md` 转好（这步轻量可自己做），建工作目录。
2. 逐阶段打开对应 stage 模板 → 填【{工作目录}】【{skill目录}】【{批次}】等占位符 → 整体作为 prompt 发给 general-purpose 子agent。
3. 每个子agent返回后**只做文件级验收**（产出文件存在、PROGRESS.md 更新、数量对），不要读全文；不达标就把验收意见+原模板重发一个新子agent返工。
4. 阶段间通过磁盘交接：题目.md → 研究契约.md / Claims-Evidence矩阵.md → 模型方案.md / 求解计划.md → all_results.json + RESULTS_q*.md + figures/ + TABLE/ → chapters/*.md → 论文.md → 评审报告_N.md → 论文.docx + check.json。禁止在对话里传大段内容。
5. ⑤评审循环里 FAIL>0 就再派一轮改稿子agent（评审者和改稿者换新的），≤3 轮，全绿才放行⑥；放行前 `scripts/freeze_numbers.py freeze` 冻结全部数字，⑥ 编译前 `check` 零漂移才许编译（防"改了代码、论文还是旧数"）。

**例外**：单点小操作（转个 PDF、跑一条命令、查个文件）不值得开子agent，主 agent 直接做。

## 单 agent 兜底模式

比赛环境若子agent不可用：仍按上述顺序执行各阶段，但每个阶段结束后把"中间结论压缩成一段话写入 PROGRESS.md 再继续"，保证任何时刻断点都能从磁盘恢复。
3. 所有数字必须可追溯（汇总进 `all_results.json`），不手编数字进论文。
4. 交付（⑥）：用本目录 `scripts/build_docx.py` 把 `论文.md` 转成符合国赛 Word 原模板格式的 .docx，再用 `scripts/validate_docx_format.py` 程序化校验（fail==0 才过），细节见 `references/14_编译与提交说明.md`：
   ```
   python scripts/build_docx.py 论文.md --out 论文.docx [--template 国赛论文模版（最强版）.docx]
   python scripts/lint_md.py 论文.md        # 转换前 md 源预检（FAIL=0 才转）
   python scripts/validate_docx_format.py 论文.docx --json check.json
   ```
   （依赖 python-docx；LaTeX 备选路线见 14 号文件的第六节，LaTeX 模板也已放在 scripts/ 下。）

## 触发示例

- "我要打数模国赛 / 华数杯，按作战包来"
- "按作战包 ② 建模" / "④ 写论文" / "⑤ 自动评审"
- "帮我改进这篇建模论文"（改进论文模式）
