#!/usr/bin/env python3
"""draw.io 逐页导出 PNG（流程图 SOP 用，见 references/25_流程图绘制SOP.md）

用法:
  python export_drawio_png.py 图.drawio -o 输出目录 [--scale 3] [--drawio "D:\\drawio\\draw.io\\draw.io.exe"]
  python export_drawio_png.py 图.drawio --list      # 只列出有几页、各页名称

说明: draw.io CLI 页码从 1 开始（v27.0.2 之前才是 0 起）；--scale 为放大倍数，
2 供检查、3 ≈ 300dpi 供论文插图。
"""
import argparse
import re
import subprocess
import sys
from pathlib import Path

DEFAULT_DRAWIO = r"D:\drawio\draw.io\draw.io.exe"


def main():
    ap = argparse.ArgumentParser(description="draw.io 逐页导出 PNG")
    ap.add_argument("src", help="输入 .drawio 文件")
    ap.add_argument("-o", "--out", default=".", help="输出目录（默认当前目录）")
    ap.add_argument("--scale", type=float, default=3.0, help="放大倍数（默认 3 ≈ 300dpi）")
    ap.add_argument("--drawio", default=DEFAULT_DRAWIO, help="draw.io.exe 路径")
    ap.add_argument("--list", action="store_true", help="只列出页数与页名")
    args = ap.parse_args()

    src = Path(args.src)
    if not src.exists():
        sys.exit(f"找不到输入文件: {src}")
    if not Path(args.drawio).exists():
        sys.exit(f"找不到 draw.io.exe: {args.drawio}（用 --drawio 指定实际路径）")

    xml = src.read_text(encoding="utf-8")
    names = re.findall(r'<diagram\b[^>]*name="([^"]*)"', xml)
    if args.list:
        for i, n in enumerate(names, 1):
            print(f"页{i}: {n}")
        return

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    fails = 0
    for p, name in enumerate(names, 1):
        dest = out / f"页{p}_{name}.png"
        r = subprocess.run(
            [args.drawio, "-x", "-f", "png", "-p", str(p),
             "-s", str(args.scale), "-o", str(dest), str(src)],
            capture_output=True, text=True,
        )
        ok = dest.exists() and dest.stat().st_size > 0
        print(("OK  " if ok else "FAIL") + f" 页{p} {name} -> {dest}")
        if not ok:
            fails += 1
            print(r.stdout + r.stderr, file=sys.stderr)
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
