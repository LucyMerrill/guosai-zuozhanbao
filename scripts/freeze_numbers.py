#!/usr/bin/env python3
"""数字冻结（防"改了代码、论文里还是旧数"，见 00 号 SOP ⑤→⑥ 交接）

用法:
  python freeze_numbers.py freeze 工作目录/all_results.json     # ⑤ 全绿后冻结
  python freeze_numbers.py check  工作目录/all_results.json     # ⑥ 编译前对账

行为:
  freeze  把 all_results.json 拍平成 键 -> 值 快照，存为同目录 frozen_numbers.json
          （含冻结时间与源文件哈希）。已有冻结文件会覆盖，覆盖前打印提示。
  check   当前 all_results.json 与冻结快照逐键比对，报 CHANGED/ADDED/REMOVED；
          有任何漂移退出码 1（即：先在 PROGRESS.md 记原因 -> 重跑受影响章节评审 -> 重新 freeze）。

纯本地比较，不联网。数值按精确相等比对（同种子重跑应当逐位一致，这正是要查的）。
"""
import hashlib
import json
import sys
from datetime import datetime
from pathlib import Path


def flatten(obj, prefix=""):
    out = {}
    if isinstance(obj, dict):
        for k, v in obj.items():
            out.update(flatten(v, f"{prefix}.{k}" if prefix else str(k)))
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            out.update(flatten(v, f"{prefix}[{i}]"))
    else:
        out[prefix] = obj
    return out


def load_flat(path: Path):
    return flatten(json.loads(path.read_text(encoding="utf-8")))


def main():
    if len(sys.argv) < 3 or sys.argv[1] not in ("freeze", "check"):
        sys.exit(__doc__)
    mode, src = sys.argv[1], Path(sys.argv[2])
    if not src.exists():
        sys.exit(f"找不到 {src}")
    frozen_path = src.parent / "frozen_numbers.json"

    if mode == "freeze":
        if frozen_path.exists():
            print(f"提示：已存在冻结文件，将覆盖（先确认漂移原因已记录在 PROGRESS.md）")
        raw = src.read_bytes()
        snap = {
            "frozen_at": datetime.now().isoformat(timespec="seconds"),
            "source": src.name,
            "source_sha256": hashlib.sha256(raw).hexdigest()[:16],
            "key_count": len(load_flat(src)),
            "numbers": load_flat(src),
        }
        frozen_path.write_text(
            json.dumps(snap, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"OK  已冻结 {snap['key_count']} 个数字 -> {frozen_path}")
        print("后续任何代码改动，⑥ 编译前必须 `check` 零漂移；有漂移先记原因再重新 freeze。")
        return

    # check
    if not frozen_path.exists():
        sys.exit(f"找不到 {frozen_path}，先执行 freeze")
    old = json.loads(frozen_path.read_text(encoding="utf-8"))["numbers"]
    new = load_flat(src)
    changed = [k for k in old if k in new and old[k] != new[k]]
    added = [k for k in new if k not in old]
    removed = [k for k in old if k not in new]
    if not (changed or added or removed):
        print(f"OK  零漂移：{len(new)} 个数字与冻结版完全一致（冻结于 {json.loads(frozen_path.read_text(encoding='utf-8'))['frozen_at']}）")
        return
    for k in changed:
        print(f"CHANGED  {k}: {old[k]!r} -> {new[k]!r}")
    for k in added:
        print(f"ADDED    {k}: {new[k]!r}")
    for k in removed:
        print(f"REMOVED  {k}")
    print(f"\n漂移 {len(changed)} 改 / {len(added)} 增 / {len(removed)} 删 —— "
          "禁止直接编译：在 PROGRESS.md 记录原因，重跑受影响章节的评审，再重新 freeze。")
    sys.exit(1)


if __name__ == "__main__":
    main()
