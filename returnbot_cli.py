#!/usr/bin/env python3
"""JSON-lines bridge between the native SwiftUI app and ReturnBot's Python core."""

import argparse
import json
import os
import queue
import sys
import threading
import uuid

from ReturnBot import ReturnBotV3, xw


RETURN_TYPES = {
    "mail-in": "Mail in",
    "mail-in-battery": "Mail in Battery",
    "kbb": "KBB",
    "kbb-battery": "KBB Battery",
}


def emit(payload):
    print(json.dumps(payload, ensure_ascii=False), flush=True)


def build_worker():
    worker = object.__new__(ReturnBotV3)
    worker.unit_price = 50.0
    worker.base_folder = os.path.dirname(os.path.abspath(__file__))
    worker.template_map = {
        "Mail in": "mail-in template.xlsx",
        "Mail in Battery": "mail-in swollen template.xlsx",
        "KBB": "kbb template.xlsx",
        "KBB Battery": "battery kbb template.xlsx",
    }
    worker.task_queue = queue.Queue()
    return worker


def preflight_excel_access():
    """Open each bundled template once so Excel requests file access up front."""
    if xw is None:
        emit({"type": "result", "operation": "preflight", "success": False,
              "message": "缺少 xlwings，無法準備 Excel。", "warnings": []})
        return 1

    worker = build_worker()
    emit({"type": "progress", "operation": "preflight", "message": "正在啟動 Excel..."})
    try:
        with xw.App(visible=False) as app:
            for template_name in dict.fromkeys(worker.template_map.values()):
                template_path = os.path.join(worker.base_folder, template_name)
                if not os.path.isfile(template_path):
                    raise FileNotFoundError(f"找不到模板：{template_name}")
                emit({
                    "type": "progress",
                    "operation": "preflight",
                    "message": f"正在確認模板存取權：{template_name}",
                })
                book = None
                try:
                    book = app.books.open(template_path, read_only=True, update_links=False)
                finally:
                    if book is not None:
                        book.close()

            emit({
                "type": "progress",
                "operation": "preflight",
                "message": "正在確認「下載項目」儲存權限...",
            })
            downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
            probe_path = os.path.join(
                downloads_path,
                f".ReturnBot_Access_Check_{uuid.uuid4().hex}.xlsx",
            )
            probe_book = None
            try:
                probe_book = app.books.add()
                probe_book.save(probe_path)
            finally:
                if probe_book is not None:
                    probe_book.close()
                if os.path.exists(probe_path):
                    os.remove(probe_path)
        emit({
            "type": "result", "operation": "preflight", "success": True,
            "message": "Excel 已就緒", "warnings": [],
        })
        return 0
    except Exception as error:
        emit({
            "type": "result", "operation": "preflight", "success": False,
            "message": f"Excel 權限預檢失敗：{error}", "warnings": [],
        })
        return 1


def main():
    parser = argparse.ArgumentParser(description="Generate ReturnBot Excel files")
    parser.add_argument("--preflight", action="store_true")
    parser.add_argument("--type", choices=RETURN_TYPES)
    parser.add_argument("--csv")
    args = parser.parse_args()

    if args.preflight:
        return preflight_excel_access()
    if not args.type or not args.csv:
        parser.error("--type 與 --csv 為生成模式的必要參數")

    csv_path = os.path.abspath(args.csv)
    if not os.path.isfile(csv_path):
        emit({"type": "result", "success": False, "message": f"找不到 CSV：{csv_path}", "warnings": []})
        return 2

    worker = build_worker()
    task = threading.Thread(
        target=worker.run_excel_task,
        args=(RETURN_TYPES[args.type], csv_path),
        daemon=False,
    )
    task.start()

    while True:
        try:
            item = worker.task_queue.get(timeout=0.1)
        except queue.Empty:
            if not task.is_alive():
                emit({"type": "result", "success": False, "message": "Python 任務結束但未回傳結果。", "warnings": []})
                return 1
            continue

        if item[0] == "status":
            emit({"type": "progress", "message": item[1]})
            continue

        if item[0] == "result":
            success, message, warnings = item[1], item[2], item[3]
            output_path = message.splitlines()[0] if success else None
            emit({
                "type": "result",
                "operation": "generation",
                "success": success,
                "message": message,
                "outputPath": output_path,
                "warnings": warnings,
            })
            task.join(timeout=1)
            return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
