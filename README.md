# TaiSan-skill 内蒙古泰三检验检测报告批量提取

批量处理 **内蒙古泰三科技服务有限公司** 出具的《检验检测报告》PDF（报告编号形如 `BG26xxxx`，封面/页眉编号 `TSKJ-JL-165`），从每份报告的封面页与信息表页提取固定字段，追加写入台账 Excel。

适用于每月例行汇总 **矿用锚索、无纵肋金属锚杆、树脂锚固剂、金属焊接网** 等煤矿物资检验检测报告的场景。可作为 **豆包 / 豆包 Turbo 等 AI 助手的技能（Skill）** 安装复用，也可作为独立 Python 工具链运行。

## 目录结构

```
TaiSan-skill/
├── SKILL.md                    # AI 助手技能定义（触发说明与执行流程）
├── references/
│   └── field_mapping.md        # PDF → Excel 字段映射规则与易错点
├── scripts/
│   ├── render_pages.py         # 把 PDF 封面页 + 信息表页渲染成 PNG
│   └── append_rows.py          # 把视觉读取的记录追加进台账 xlsx（自动续号/反查采样地点）
├── examples/
│   └── rows.example.json       # rows.json 输入示例
├── requirements.txt
├── LICENSE                     # MIT
└── .gitignore
```

## 安装依赖

```bash
pip install -r requirements.txt
```

需要 `openpyxl` 与 `PyMuPDF`。

## 使用流程

### 1. 渲染关键页为 PNG

泰三 PDF 是扫描件，基本没有文本层，`get_text()` 取不到标签，必须先渲染成图片再视觉读取：

```bash
python scripts/render_pages.py --src "报告目录" --out "临时输出目录"
```

每份 PDF 会输出封面页与信息表页 PNG。找不到信息表页时回退到第 3 页。

### 2. 视觉读取并拼 rows.json

按 `references/field_mapping.md` 的映射规则，从信息表页 PNG 抄字段，写成 `rows.json`（格式见 `examples/rows.example.json`）。

要点：
- **通知检测单位** = 报告"抽（送）样地点"的**公司全称**；
- **采样地点** 不要从报告抄，留空，脚本会按"通知检测单位"自动反查映射表填入；
- 找不到的字段留空字符串并在"备注"注明"未找到"，绝不编造；
- 数量、规格型号原样照录（含 ×、Φ、/、空格、单位），不换算。

### 3. 追加写入台账

```bash
python scripts/append_rows.py --xlsx "台账.xlsx" --rows rows.json
```

脚本行为：
- 按表头第 1 行自动对齐列，无需硬编码列号；
- "序号"列自动在上一最大值上自增；
- 写前自动备份一份 `<文件名>.备份.xlsx`；
- 按"通知检测单位"自动反查填充"采样地点"（内置映射表，见 `field_mapping.md`）；
- 沿用既有行样式。

### 4.（可选）使用外部采样地点映射表

内置了内蒙古仲泰系各矿的《公司全称 → 采样地点》映射表。若需适配其他检测单位，可传入外部 JSON 覆盖/补充：

```bash
python scripts/append_rows.py --xlsx "台账.xlsx" --rows rows.json --station-map station_map.json
```

`station_map.json` 格式：

```json
{
  "内蒙古某检测有限公司": "某采样地点"
}
```

外部条目覆盖同键内置条目，其余仍走内置映射表。

## 硬约束（录入口径）

- 找不到的字段留空 + 备注"未找到"，绝不编造数值/日期/编号。
- 数量、规格型号原样照录（含单位与符号），不四舍五入、不换算单位。
- 取样时间统一 `YYYY/MM/DD`，按文本存储（避免 Excel 自动改格式）。
- 结论只写"合格/不合格"：结论原文"所检项目符合……技术要求"即判合格。
- 若保存报 `PermissionError`，说明 xlsx 正被 Excel 占用，请先关闭 Excel 再重跑。

## 作为 AI 助手技能安装

把本仓库克隆或下载后，将整个目录（含 `SKILL.md`）放入助手技能根目录即可被识别。当用户提到"处理泰三检验报告""批量提取泰三检测报告""内蒙古泰三 PDF 出表"等需求时，助手会按 `SKILL.md` 流程执行。

## 许可证

[MIT](LICENSE)
