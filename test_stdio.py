#!/usr/bin/env python3
"""stdio 真实验收：官方 Client SDK 握手 -> list_tools -> 4 个工具真实调用。"""
import asyncio
import csv
import json
import os
import sys

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

HERE = os.path.dirname(os.path.abspath(__file__))
SERVER = os.path.join(HERE, "src", "excel_toolkit_mcp", "server.py")
SCRATCH = "/root/.hermes/cache/scratch/mcp_pkg_test"
os.makedirs(SCRATCH, exist_ok=True)
INPUT_CSV = os.path.join(SCRATCH, "excel_messy.csv")
SPEC_JSON = os.path.join(SCRATCH, "excel_clean.json")
OUTPUT_XLSX = os.path.join(SCRATCH, "excel_cleaned.xlsx")

results = {}


def _prep():
    with open(INPUT_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([" 姓名（Name） ", "金额", "日期", "地区", "空列"])
        w.writerow([" 张三 ", "￥1,200.50", "2026年9月1日", "郑州,河南", ""])
        w.writerow(["张三", "1200.5", "2026-09-01", "郑州,河南", ""])
        w.writerow(["李四", "$99", "2026.09.02", "北京,北京", ""])
        w.writerow(["王五", "0元", "2026/9/3", "上海,上海", ""])
    spec = {
        "normalize_headers": True,
        "drop_empty_cols": True,
        "drop_duplicates": True,
        "strip_whitespace": True,
        "number_columns": ["金额"],
        "date_columns": ["日期"],
        "split_columns": [{"column": "地区", "into": ["市", "省"], "separator": ","}],
        "select_columns": ["姓名", "金额", "日期", "市", "省"],
        "filter": {"column": "金额", "op": "gt", "value": 0},
    }
    with open(SPEC_JSON, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False)


async def main():
    _prep()
    params = StdioServerParameters(command=sys.executable, args=[SERVER])
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            init = await session.initialize()
            results["initialize"] = {
                "serverName": init.server_info.name,
                "serverVersion": init.server_info.version,
                "protocol": init.protocol_version,
            }

            tools = await session.list_tools()
            names = [t.name for t in tools.tools]
            results["list_tools"] = names
            assert set(names) == {
                "generate_formula", "list_formula_intents", "clean_table", "selftest"
            }, names

            # 1) 公式：确定性格局
            r = await session.call_tool(
                "generate_formula",
                {"text": "计算B2到B100的平均值，忽略空单元格"},
            )
            f1 = json.loads(r.content[0].text)
            assert f1["formula"] == "=AVERAGE(B2:B100)", f1

            r = await session.call_tool(
                "generate_formula", {"text": "统计A列满足大于60的有几个"}
            )
            f2 = json.loads(r.content[0].text)
            assert f2["formula"].startswith("=COUNTIF"), f2

            r = await session.call_tool(
                "generate_formula", {"text": "聊聊天气吧"}
            )
            f3 = json.loads(r.content[0].text)
            assert f3["formula"] is None and "hint" in f3, f3
            results["formula"] = {
                "average": f1["formula"],
                "countif": f2["formula"],
                "unknown_returns_null": True,
            }

            # 2) 意图清单
            r = await session.call_tool("list_formula_intents", {})
            intents = json.loads(r.content[0].text)
            assert intents["count"] >= 25, intents["count"]
            results["intent_count"] = intents["count"]

            # 3) 清洗：真实落盘
            if os.path.exists(OUTPUT_XLSX):
                os.remove(OUTPUT_XLSX)
            r = await session.call_tool(
                "clean_table",
                {
                    "input_path": INPUT_CSV,
                    "spec": json.load(open(SPEC_JSON, encoding="utf-8")),
                    "output_path": OUTPUT_XLSX,
                },
            )
            cleaned = json.loads(r.content[0].text)
            assert cleaned["ok"] and os.path.isfile(OUTPUT_XLSX), cleaned
            assert cleaned["columns"] == ["姓名", "金额", "日期", "市", "省"], cleaned
            # 张三重复行去重 + 王五金额=0 被过滤 -> 表头 + 张三 + 李四 = 3 行
            assert cleaned["rows_including_header"] == 3, cleaned
            results["clean"] = {
                "output_exists": True,
                "size": cleaned["output_size_bytes"],
                "rows": cleaned["rows_including_header"],
                "columns": cleaned["columns"],
                "notes": cleaned["notes"],
            }

            # 用 openpyxl 复核落盘内容（日期归一/金额归一/拆列）
            import openpyxl
            wb = openpyxl.load_workbook(OUTPUT_XLSX)
            ws = wb.active
            data = [list(row) for row in ws.iter_rows(values_only=True)]
            assert data[1] == ["张三", 1200.5, "2026-09-01", "郑州", "河南"], data
            assert data[2][0] == "李四" and data[2][1] == 99 and data[2][2] == "2026-09-02", data
            results["clean"]["disk_verify"] = data

            # 4) 错误路径
            r = await session.call_tool(
                "clean_table",
                {"input_path": "/nonexistent/x.csv", "spec": {}, "output_path": "/tmp/x.xlsx"},
            )
            assert bool(r.is_error), "不存在文件应 is_error"
            results["error_handling"] = {"is_error": True}

            # 5) 自检
            r = await session.call_tool("selftest", {})
            st = json.loads(r.content[0].text)
            assert st["ok"] is True, st
            results["selftest"] = "pass"

            print(json.dumps(results, ensure_ascii=False, indent=2, default=str))
            print("\nALL ASSERTIONS PASSED")


asyncio.run(main())
