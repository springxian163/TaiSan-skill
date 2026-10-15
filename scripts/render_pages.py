# -*- coding: utf-8 -*-
"""渲染内蒙古安创检验报告 PDF 的关键页为 PNG，供视觉读取。

每份报告结构固定：
  - 第1页：封面（左上角报告编号；底部为检测单位名称）
  - "报告正文第1页"：双列信息表（委托单位/委托人/样品名称/规格型号/抽(送)样日期/
    抽(送)样地点/抽(送)样人员/生产单位/样品数量/检验结论 等）。该页文本层标签与数值
    会被打乱，必须渲染成图片后用视觉读取，不要只靠 get_text()。

用法:
  python render_pages.py --src "D:\\测试AI\\报告" --out "<某临时目录>"

输出: 每份 PDF 的 封面页 与 信息表页 PNG，并打印路径清单。
"""
import argparse, os, glob, sys
import pymupdf


def find_info_page(doc):
    """信息表页 = 同时含 '抽（送）样日期' 与 '检验结论' 的页；找不到回退第3页(index2)。"""
    for i, page in enumerate(doc):
        t = page.get_text()
        if ("抽（送）样日期" in t or "抽(送)样日期" in t) and "检验结论" in t:
            return i
    return 2 if doc.page_count >= 3 else 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True, help="PDF 所在目录")
    ap.add_argument("--out", required=True, help="PNG 输出目录")
    ap.add_argument("--zoom", type=float, default=2.2)
    a = ap.parse_args()

    os.makedirs(a.out, exist_ok=True)
    pdfs = sorted(glob.glob(os.path.join(a.src, "*.pdf")))
    if not pdfs:
        print("未找到 PDF: " + a.src); sys.exit(1)

    for f in pdfs:
        doc = pymupdf.open(f)
        base = os.path.splitext(os.path.basename(f))[0].replace("/", "_")[:40]
        targets = [0, find_info_page(doc)]
        targets = sorted(set(targets))
        for i in targets:
            if i >= doc.page_count:
                continue
            pix = doc[i].get_pixmap(matrix=pymupdf.Matrix(a.zoom, a.zoom), alpha=False)
            fn = os.path.join(a.out, f"{base}_p{i+1}.png")
            pix.save(fn)
            kind = "封面" if i == 0 else "信息表"
            print(f"{kind}\t{fn}")
        doc.close()


if __name__ == "__main__":
    main()
