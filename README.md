# excel-toolkit MCP

纯本地 Excel/CSV 工具的 MCP 封装（无 API Key、无 LLM、不联网）。核心引擎来自
`data/disk/projects/excel-toolkit`（确定性公式模板 + 表格清洗）。

## 工具

| 工具 | 说明 |
|---|---|
| `generate_formula(text)` | 一句中文 → 可直接粘贴的 Excel/Google Sheets 公式（30+ 意图，识别不了返回 `formula: null`） |
| `list_formula_intents()` | 列出全部支持的意图 |
| `clean_table(input_path, spec, output_path, as_csv=false)` | CSV/XLSX 按 JSON 指令清洗并落盘 |
| `selftest()` | 核心引擎自检 |

`spec` 支持：`normalize_headers / drop_empty_cols / drop_duplicates /
strip_whitespace / number_columns / date_columns / split_columns /
select_columns / fillna / filter / dedup_subset`。

## 运行

```bash
pip install -r requirements.txt
python server.py                      # stdio
python server.py --transport streamable-http --port 8766
```

## 实测

```bash
python test_stdio.py
```
