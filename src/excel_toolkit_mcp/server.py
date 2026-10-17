#!/usr/bin/env python3
"""
excel-toolkit MCP Server — 纯本地 Excel 公式生成 + CSV/XLSX 清洗

暴露工具：
  generate_formula(text)                        -> {formula, intent/hint}
  list_formula_intents()                        -> 支持的中文意图列表
  clean_table(input_path, spec, output_path, csv=False) -> 清洗落盘结果
  selftest()                                    -> 核心引擎自检

基于 excel-toolkit skill（excel_core.py，确定性引擎，无 LLM、无网络）。
传输：stdio（默认）/ streamable-http（--transport streamable-http --port N）
兼容 mcp SDK 1.x（FastMCP）与 2.x（MCPServer）双导入。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from . import excel_core as core
except ImportError:
    import excel_core as core

try:  # mcp 2.x
    from mcp.server.mcpserver import MCPServer
except ImportError:  # mcp 1.x
    from mcp.server.fastmcp import FastMCP as MCPServer

mcp = MCPServer(
    name="excel-toolkit",
    instructions=(
        "纯本地 Excel/CSV 工具。generate_formula 用一句中文返回可直接粘贴的 "
        "Excel/Google Sheets 公式；clean_table 按 JSON 清洗指令把 CSV/XLSX "
        "洗成规范 XLSX/CSV（去重/去空格/日期金额归一/拆列/筛选等）。"
        "所有路径必须是服务器本地绝对路径，不联网。"
    ),
)


@mcp.tool(
    name="generate_formula",
    description=(
        "用一句中文生成可直接粘贴的 Excel/Google Sheets 公式（确定性模板引擎，"
        "非大模型）。覆盖 SUM/AVERAGE/COUNTIF/SUMIF/IF/VLOOKUP/XLOOKUP/"
        "TEXTJOIN/DATEDIF/UNIQUE/RANK/占比/增长率等 30+ 高频意图。"
        "识别不了时 formula 返回 null 并给 hint。"
    ),
)
def generate_formula(text: str) -> dict:
    if not text or not text.strip():
        raise ValueError("text 不能为空")
    return core.formula(text)


@mcp.tool(
    name="list_formula_intents",
    description="列出公式生成器支持的全部中文意图说明。",
)
def list_formula_intents() -> dict:
    return {
        "count": len(core.INTENTS),
        "intents": [
            {"pattern": pat, "description": desc} for pat, _, desc in core.INTENTS
        ],
    }


@mcp.tool(
    name="clean_table",
    description=(
        "按 JSON 清洗指令清洗本地 CSV/XLSX 并落盘。spec 支持键："
        "normalize_headers/drop_empty_cols/drop_duplicates/strip_whitespace/"
        "number_columns/date_columns/split_columns/select_columns/fillna/filter/"
        "dedup_subset。CSV 编码自动尝试 utf-8-sig/gbk。"
        "返回输出路径、行数（含表头）与清洗动作日志。"
    ),
)
def clean_table(
    input_path: str,
    spec: dict,
    output_path: str,
    as_csv: bool = False,
) -> dict:
    inp = os.path.abspath(os.path.expanduser(input_path))
    if not os.path.isfile(inp):
        raise FileNotFoundError(f"输入文件不存在: {inp}")
    if not isinstance(spec, dict):
        raise ValueError("spec 必须是 JSON 对象")
    out = os.path.abspath(os.path.expanduser(output_path))
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    if as_csv and not out.lower().endswith(".csv"):
        out = os.path.splitext(out)[0] + ".csv"
    if not as_csv and not out.lower().endswith((".xlsx", ".csv")):
        out += ".xlsx"

    rows = core.read_table(inp)
    cleaned, notes = core.clean_rows(rows, spec)
    core.write_table(cleaned, out, as_csv or out.lower().endswith(".csv"))
    return {
        "ok": True,
        "input_path": inp,
        "output_path": out,
        "output_size_bytes": os.path.getsize(out),
        "rows_including_header": len(cleaned),
        "columns": cleaned[0] if cleaned else [],
        "notes": notes,
    }


@mcp.tool(
    name="selftest",
    description="运行核心引擎自检（公式顺序/条件意图、日期三种格式、清洗拆列对齐等）。",
)
def selftest() -> dict:
    import contextlib
    import io
    buf = io.StringIO()
    try:
        # _selftest 会往 stdout 打印 PASS 行，stdio 传输下必须吞掉
        with contextlib.redirect_stdout(buf):
            core._selftest()
        return {"ok": True, "detail": buf.getvalue().strip()}
    except AssertionError as e:
        return {"ok": False, "error": str(e)}


def main():
    import argparse

    ap = argparse.ArgumentParser(description="excel-toolkit MCP server")
    ap.add_argument("--transport", default="stdio",
                    choices=["stdio", "streamable-http", "sse"])
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=8766)
    args = ap.parse_args()

    if args.transport == "stdio":
        mcp.run(transport="stdio")
    else:
        mcp.run(transport=args.transport, host=args.host, port=args.port)


if __name__ == "__main__":
    main()
