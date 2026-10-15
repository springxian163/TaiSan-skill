# -*- coding: utf-8 -*-
"""把从 PDF 视觉读出的记录追加进台账 xlsx。

列以表头行(第1行)为准自动对齐，无需硬编码列号；"序号"列自动在上一最大值上自增。
写前自动备份一份 <文件名>.备份.xlsx。若目标文件被 Excel 占用会报 PermissionError，
提示先关闭 Excel 再重跑。

采样地点自动反查：
  脚本内置《通知检测单位 → 采样地点》映射表，也可通过 --station-map 传入外部 JSON
  覆盖/补充。若 rows.json 里某行未提供"采样地点"列（或留空），会自动根据
  "通知检测单位"列的值反查映射表并写入"采样地点"列。命中失败则"采样地点"留空，
  并在该行"备注"末尾追加"采样地点未匹配：<原公司名>"，绝不编造。

用法:
  python append_rows.py --xlsx "台账.xlsx" --rows rows.json
  python append_rows.py --xlsx "台账.xlsx" --rows rows.json --station-map station_map.json

rows.json 形如:
[
  {
    "取样人员": "赵忠全、郭常春",
    "供应站业务员/委托人": "丁健平",
    "样品供货商": "伊金霍洛旗普惠工贸有限公司",
    "型号及品牌": "矿用锚索 SKL1×19S-21.8/1860 Φ21.8×6300",
    "抽样数量": "3根、4块、4个",
    "取样时间": "2026/07/28",
    "报告编号": "BG260078",
    "检测报告状况": "合格",
    "通知检测单位": "内蒙古仲泰能源集团有限公司伊旗分公司",
    "检测单位": "内蒙古泰三科技服务有限公司",
    "备注": ""
  }
]
未提供的列留空。"采样地点"若留空或不填，脚本会按"通知检测单位"自动反查。

station_map.json 形如:
{"内蒙古XX检测有限公司": "某采样地点"}
"""
import argparse, json, os, re, shutil, sys
from copy import copy
import openpyxl

SEQ_HEADER = "序号"
STATION_HEADER = "采样地点"
LAB_HEADER = "通知检测单位"

# 通知检测单位(公司全称) → 采样地点(煤矿名)
# 来源: 用户 2026-10-09 提供的对应表; 方家畔煤矿对应两家公司均保留。
COMPANY_TO_MINE = {
    "内蒙古仲泰能源集团有限公司伊旗分公司": "红庆河煤矿",
    "内蒙古仲泰能源集团有限公司纳林陶亥分公司": "大地精煤矿",
    "内蒙古仲泰能源集团有限公司羊市塔分公司": "凯达煤矿",
    "内蒙古仲泰酸刺沟矿业有限公司": "酸刺沟煤矿",
    "内蒙古仲泰塔拉壕矿业有限公司": "塔拉壕煤矿",
    "鄂尔多斯市源泰矿业有限责任公司": "宝山煤矿",
    "内蒙古仲泰能源集团有限公司宏景塔分公司": "宏一煤矿",
    "内蒙古仲泰能源集团有限公司准旗分公司": "黑岱沟煤矿",
    "内蒙古仲泰闫家沟矿业有限公司": "闫家沟煤矿",
    "榆林市榆阳区方家畔矿业有限责任公司": "方家畔煤矿",
    "内蒙古仲泰能源集团有限公司神木分公司": "朱盖塔煤矿",
    "鄂尔多斯市嘉泰矿业有限责任公司榆林分公司": "方家畔煤矿",
}


def _norm(s):
    """去空白、统一全半角，用于匹配公司名。"""
    if s is None:
        return ""
    s = str(s).strip()
    # 全角字母数字转半角
    out = []
    for ch in s:
        code = ord(ch)
        if 0xFF01 <= code <= 0xFF5E:
            out.append(chr(code - 0xFEE0))
        elif code == 0x3000:
            out.append(" ")
        else:
            out.append(ch)
    return re.sub(r"\s+", "", "".join(out))


# 预生成归一化后的映射
_NORM_MAP = {_norm(k): v for k, v in COMPANY_TO_MINE.items()}


def load_external_map(path):
    """从外部 JSON 文件加载《通知检测单位 → 采样地点》映射，并合入内置表。
    JSON 形如: {"内蒙古XX有限公司": "某采样地点"}，外部条目覆盖同键内置条目。"""
    global COMPANY_TO_MINE, _NORM_MAP
    with open(path, encoding="utf-8-sig") as fh:
        ext = json.load(fh)
    if not isinstance(ext, dict):
        raise ValueError("station-map JSON 顶层必须是对象(公司全称 -> 采样地点)")
    COMPANY_TO_MINE.update({str(k): str(v) for k, v in ext.items()})
    _NORM_MAP = {_norm(k): v for k, v in COMPANY_TO_MINE.items()}


def lookup_station(lab_unit):
    """根据通知检测单位反查采样地点。返回 (采样地点 or None, 原公司名)。"""
    key = _norm(lab_unit)
    if not key:
        return None, ""
    return _NORM_MAP.get(key), str(lab_unit).strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--rows", required=True)
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--station-map", default=None,
                    help="可选：外部《通知检测单位→采样地点》映射 JSON，覆盖/补充内置映射表")
    a = ap.parse_args()

    if a.station_map and os.path.exists(a.station_map):
        load_external_map(a.station_map)
        print(f"已加载外部映射表: {a.station_map}")

    if not os.path.exists(a.xlsx):
        print("找不到 xlsx: " + a.xlsx); sys.exit(1)

    backup = os.path.join(os.path.dirname(a.xlsx),
                          os.path.splitext(os.path.basename(a.xlsx))[0] + ".备份.xlsx")
    shutil.copyfile(a.xlsx, backup)
    print("备份 ->", backup)

    with open(a.rows, encoding="utf-8-sig") as fh:
        rows = json.load(fh)

    # 写入前先做采样地点自动反查
    resolved, missed = 0, 0
    for rec in rows:
        station = (rec.get(STATION_HEADER) or "").strip()
        if station:
            continue  # 已显式提供则不覆盖
        lab = (rec.get(LAB_HEADER) or "").strip()
        mine, raw = lookup_station(lab)
        if mine:
            rec[STATION_HEADER] = mine
            resolved += 1
        else:
            rec[STATION_HEADER] = ""
            missed += 1
            note = (rec.get("备注") or "").strip()
            tag = f"采样地点未匹配：{raw}" if raw else "采样地点未匹配：通知检测单位为空"
            rec["备注"] = (note + "；" + tag) if note else tag
    print(f"采样地点反查: 命中 {resolved}, 未命中 {missed}")

    wb = openpyxl.load_workbook(a.xlsx)
    ws = wb[a.sheet] if a.sheet else wb.worksheets[0]

    # 表头 -> 列号
    header_col = {}
    for c in ws[1]:
        if c.value is not None:
            header_col[str(c.value).strip()] = c.column

    # 现有最大序号
    max_seq = 0
    seq_col = header_col.get(SEQ_HEADER)
    if seq_col:
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=seq_col).value
            if isinstance(v, (int, float)):
                max_seq = max(max_seq, int(v))

    # 样式模板：优先用上一数据行，否则用表头行
    tmpl_row = ws.max_row if ws.max_row >= 2 else 1

    start = ws.max_row + 1
    for i, rec in enumerate(rows):
        r = start + i
        max_seq += 1
        if seq_col:
            ws.cell(row=r, column=seq_col, value=max_seq)
        for k, v in rec.items():
            k = str(k).strip()
            if k == SEQ_HEADER or k not in header_col:
                continue
            cell = ws.cell(row=r, column=header_col[k], value=v)
            tmpl = ws.cell(row=tmpl_row, column=header_col[k])
            cell.font = copy(tmpl.font); cell.border = copy(tmpl.border)
            cell.alignment = copy(tmpl.alignment); cell.fill = copy(tmpl.fill)
        if seq_col:
            c = ws.cell(row=r, column=seq_col)
            tmpl = ws.cell(row=tmpl_row, column=seq_col)
            c.font = copy(tmpl.font); c.border = copy(tmpl.border)
            c.alignment = copy(tmpl.alignment); c.fill = copy(tmpl.fill)

    wb.save(a.xlsx)
    print(f"已追加 {len(rows)} 行，末行={start+len(rows)-1}，下一序号={max_seq+1}")


if __name__ == "__main__":
    main()
