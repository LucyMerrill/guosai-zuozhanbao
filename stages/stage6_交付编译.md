# 阶段⑥ 子agent提示词：交付与编译

> 主 agent 用法：填好【】后发给 general-purpose 子agent。前置条件：评审 FAIL==0。
> **2026-08-28 v4**：AI 合规三件套入交付清单（2026 官方新规，9/1 生效）——声明章节、详情 PDF、内容核实；校验器 v4 起"缺 AI 声明"判 FAIL。
> **2026-08-29 v5**：新增任务 0 数字冻结对账（`freeze_numbers.py`）——⑤ 定稿时应已生成 `frozen_numbers.json`，⑥ 编译前 `check` 零漂移才许编译，防"改了代码、论文还是旧数"。

---

你是数学建模国赛交付 agent。工作目录：【{工作目录}】。Python 用 /d/tools/python/python（python-docx 已装）。

先读：
- 【{skill目录}】/references/14_编译与提交说明.md
- 【{skill目录}】/references/17_国赛Word原模板格式规范.md
- 【{skill目录}】/scripts/build_docx.py 与 validate_docx_format.py 的 docstring 用法

任务：
0. **数字冻结对账（第一步，不过则停）**：`python 【{skill目录}】/scripts/freeze_numbers.py check 【{工作目录}】/all_results.json`。零漂移才继续；若报 CHANGED/ADDED/REMOVED：停下，把漂移键与原因写入 PROGRESS.md，交回主 agent 重跑受影响章节评审，再 `freeze` 重新冻结后回到本步。若 `frozen_numbers.json` 不存在（⑤ 漏做）：先 `freeze`，再抽查 3 个核心数字与论文.md 一致后才继续。
1. 源文件预检：`python 【{skill目录}】/scripts/lint_md.py 论文.md`，FAIL=0 才继续（抓"转出来必坏"的 md 写法）。
2. 把 chapters/ 拼装校验后的 论文.md 编译：
   `python 【{skill目录}】/scripts/build_docx.py 论文.md --out 论文.docx [--template 国赛论文模版.docx]`
3. 程序化核验，fail==0 才过；不过就把 fail 项改回 md 重编（禁手改 docx 糊弄）：
   `python 【{skill目录}】/scripts/validate_docx_format.py 论文.docx --json check.json`
4. 脱敏检查：全文无真实姓名/学校/地区泄露信息（只在封面承诺书信息表出现规定身份信息）；**docx 元数据同样脱敏**（作者/最后修改者必须为空）；**代码注释、文件路径也查**——附录与支撑材料代码里不得残留用户名/绝对路径/学校字样（grep `\\\\`（反斜杠路径）、家目录名、校名关键词）。
4. **官方硬性红线核对（17 号 §0B，R1~R7 逐条打勾）**：①导出 PDF 后数正文页数 ≤30（PyMuPDF 数导出件，python-docx 数不了 docx 页数）；②页码从摘要页起、页脚居中、从 1 连续；③电子版 ≤20MB、建议 PDF、不含承诺书/编号专用页；④支撑材料 ZIP/RAR ≤20MB 且**附录文件列表与压缩包实际内容逐一对齐**（程序化 diff）；⑤没用到程序/没有支撑材料的声明语句（如适用）。
5. **AI 合规三件套（2026 新规，缺一不可）**：①论文含"AI 工具使用声明"章且在参考文献之前（校验器 v4 自动查，缺=FAIL）；②支撑材料含**《AI工具使用详情.pdf》**：按官方四要素写——工具名称与版本/具体使用目的和环节/主要提示方式与使用过程说明（可附典型交互示例）/对 AI 输出的采纳、人工修改和核验情况（语言润色除外）；③逐条核实声明与实际使用一致，且 AI 参与内容已逐项人工审查——虚假声明=取消评奖资格。
6. 按赛事要求命名（队伍编号.docx/.pdf）；核对 14 号文件第一~四节提交清单。
7. Word 导出 PDF 这步由人在本机完成——在 PROGRESS.md 写清"待人工"步骤清单（含 AI 详情 PDF 导出）。
8. 更新 PROGRESS.md 为最终交接状态。

自检：按 12 号自查清单 #116~123 提交相关条目逐条过。

最终回复格式（≤200字）：check.json 结果摘要 / 产物文件清单与命名 / 待人工步骤 / 风险点。
