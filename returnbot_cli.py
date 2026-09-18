#!/usr/bin/env python3
"""JSON-lines bridge between the native SwiftUI app and ReturnBot's Python core."""

import argparse
import json
import os
import queue
import sys
import threading


RETURN_TYPES = {
    "mail-in": "Mail in",
    "mail-in-battery": "Mail in Battery",
    "kbb": "KBB",
    "kbb-battery": "KBB Battery",
}


def emit(payload):
    print(json.dumps(payload, ensure_ascii=False), flush=True)


def build_worker():
    from ReturnBot import ReturnBotV3
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


def main():
    parser = argparse.ArgumentParser(description="Generate ReturnBot Excel files")
    parser.add_argument("--type", choices=RETURN_TYPES)
    parser.add_argument("--csv")
    parser.add_argument("--output-directory", help="Optional output folder; defaults to Downloads")
    parser.add_argument("--recall", choices=['prices', 'preview', 'export'])
    args = parser.parse_args()

    if args.recall:
        from recall import handle
        try:
            result = handle(args.recall, json.load(sys.stdin))
            emit({'success': True, 'data': result})
            return 0
        except Exception as error:
            emit({'success': False, 'message': str(error)})
            return 1

    if not args.type or not args.csv:
        parser.error("--type 與 --csv 為生成模式的必要參數")

    csv_path = os.path.abspath(args.csv)
    if not os.path.isfile(csv_path):
        emit({"type": "result", "success": False, "message": f"找不到 CSV：{csv_path}", "warnings": []})
        return 2

    worker = build_worker()
    if args.output_directory:
        worker.output_directory = os.path.abspath(args.output_directory)
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
