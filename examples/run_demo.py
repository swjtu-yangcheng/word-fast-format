# -*- coding: utf-8 -*-
"""
一键示例：Markdown → docx → 格式体检 → 套预设排版。

用法：
    python examples/run_demo.py

产出：
    examples/out/demo_raw.docx     直接转换（未排版）
    examples/out/demo_formatted.docx  套用 cn-formal 预设后
"""
import os
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
WF = os.path.join(ROOT, "scripts", "wordfmt.py")
OUT = os.path.join(HERE, "out")


def run(args):
    return subprocess.run([sys.executable, WF] + args,
                          capture_output=True, text=True, encoding="utf-8")


def main():
    os.makedirs(OUT, exist_ok=True)
    src = os.path.join(HERE, "sample.md")
    raw = os.path.join(OUT, "demo_raw.docx")
    fmt = os.path.join(OUT, "demo_formatted.docx")

    # 刻意先用 academic（期刊论文：宋体五号）生成，再一键换成 cn-formal（公文规范），
    # 以便直观对比"整篇换排版"的效果
    t0 = time.time()
    r = run(["md2docx", src, "--preset", "academic", "--out", raw])
    print(r.stdout or r.stderr)
    print(f"[1/4] Markdown → docx（按 academic 论文样式）  用时 {time.time() - t0:.2f}s")

    print("\n[2/4] 排版前体检：")
    r = run(["inspect", raw])
    print(r.stdout or r.stderr)

    t0 = time.time()
    r = run(["apply", raw, "--preset", "cn-formal", "--out", fmt, "--report"])
    print(r.stdout or r.stderr)
    print(f"[3/4] 一键换成 cn-formal 公文规范  用时 {time.time() - t0:.2f}s  → {fmt}")

    print("[4/4] 排版后体检：")
    r = run(["inspect", fmt])
    print(r.stdout or r.stderr)

    print(f"\n完成。对比查看：\n  {raw}\n  {fmt}")


if __name__ == "__main__":
    main()
