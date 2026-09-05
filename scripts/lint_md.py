# -*- coding: utf-8 -*-
"""
国赛 md 源文件预检器（build_docx.py 的上游闸门）
================================================
docx 成品格式由 validate_docx_format.py 把关；本脚本管**源头 论文.md**，

  1. 裸 LaTeX 命令    —— 不在 $$..$$ / $..$ / 代码块 里的 \\frac 等会被当字面文本排进 Word
  2. 代码围栏不闭合   —— ``` 数量为奇数，后半篇全被吞成代码
  3. 伪标题           —— "- 一、xxx" 列表行冒充标题（转出来丢层级）
  4. 伪图注/伪表注    —— "- 图 3：xxx" 列表行冒充题注（应为独立一行"图 3 xxx"）
  5. 公式中文括号编号 —— $$..$$（12）→ build_docx 只认半角 (12)
  6. $$ 行混入正文    —— 公式行公式外还有文字（除编号 (n)/（n)），转 docx 易错位

用法：python lint_md.py 论文.md         # 退出码 0=干净 1=有违规
"""
import re
import sys

LATEX_CMDS = re.compile(
    r'\\(tag|frac|tfrac|sqrt|hat|dot|ddot|vec|bar|text|mathbf|mathcal'
    r'|begin\{(?:aligned|cases|bmatrix|pmatrix|equation|array|figure|table)'
    r'|cdot|times|pm|mp|leq|geq|neq|approx|sum|prod|int|i?int|partial|nabla|infty'
    r'|left|right|big[glr]?|cap|cup|alpha|beta|gamma|theta|lambda|mu|sigma|omega)\b')

PSEUDO_TITLE = re.compile(r'^\s*[-*]\s*(第?[一二三四五六七八九十\d]+[、.．]\s*\S|摘\s*要|关键词\s*[：:]|参考文献|附录)')
PSEUDO_CAPTION = re.compile(r'^\s*[-*]\s*[图表]\s*\d+[：:]')
CN_PAREN_NUM = re.compile(r'（\d+）')


def lint(path):
    lines = open(path, encoding='utf-8').read().splitlines()
    problems = []  # (级别, 类别, 行号, 摘录)

    fence_count = 0
    in_code = False
    in_dollar = False
    for idx, ln in enumerate(lines, 1):
        stripped = ln.strip()
        if stripped.startswith('```'):
            fence_count += 1
            in_code = not in_code
            continue
        if in_code:
            continue
        safe = False  # 该行是否属于公式/代码等"豁免区"
        if not in_code:
            n = stripped.count('$$')
            if n >= 2:
                # 行内完整公式：检查公式外是否混入正文
                outside = re.sub(r'\$\$[^$]+\$\$', '', stripped).strip()
                outside_stripped_num = re.sub(r'^[（(]\d+[)）]\s*$', '', outside)  # 行尾编号豁免
                if CN_PAREN_NUM.search(outside):
                    problems.append(('WARN', '公式中文括号编号(应为半角)', idx, stripped[:70]))
                if outside_stripped_num:
                    problems.append(('WARN', '$$行混入正文', idx, stripped[:70]))
                safe = True
            elif n == 1:
                in_dollar = not in_dollar
                safe = True
                if CN_PAREN_NUM.search(ln):
                    problems.append(('WARN', '公式中文括号编号', idx, stripped[:70]))
        if in_dollar:
            safe = True
            continue
        if safe:
            continue

        # ---- 以下检查作用于"普通正文行" ----
        probe = re.sub(r'\$[^\n$]+\$', '', ln)  # 去行内公式后扫描
        if LATEX_CMDS.search(probe):
            problems.append(('FAIL', '裸LaTeX命令(缺$$包围)', idx, stripped[:70]))
        if PSEUDO_TITLE.match(ln):
            problems.append(('FAIL', '伪标题(列表冒充章节)', idx, stripped[:70]))
        if PSEUDO_CAPTION.match(ln):
            problems.append(('FAIL', '伪图注/伪表注(列表冒充题注)', idx, stripped[:70]))

    if fence_count % 2 == 1:
        problems.append(('FAIL', '代码围栏不闭合', 0, f'``` 共 {fence_count} 个（奇数），后半篇会被吞进代码块'))

    return problems


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    problems = lint(sys.argv[1])
    if not problems:
        print(f'✓ {sys.argv[1]} 预检通过（无"转出来必坏"写法）')
        sys.exit(0)
    fails = [p for p in problems if p[0] == 'FAIL']
    warns = [p for p in problems if p[0] == 'WARN']
    for lv, cat, idx, excerpt in problems:
        loc = f'L{idx}' if idx else '--'
        print(f'[{lv}] {cat} @{loc}: {excerpt}')
    print(f'—— 共 {len(fails)} FAIL / {len(warns)} WARN，FAIL 必须修完再转 docx ——')
    sys.exit(1 if fails else 0)


if __name__ == '__main__':
    main()
