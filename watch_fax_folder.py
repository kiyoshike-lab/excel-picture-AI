from pathlib import Path
from time import sleep

from app import INCOMING, process_incoming_once


def main():
    INCOMING.mkdir(exist_ok=True)
    print(f"受信FAXフォルダを監視中: {INCOMING}")
    print("複合機のFAX保存先をこのフォルダにすると、自動でExcelを作成します。")
    while True:
        results = process_incoming_once()
        for result in results:
            if result["status"] == "created":
                print(f"Excel作成: {result['source']} -> {result['file']}")
            else:
                print(f"確認が必要: {result['source']} / {result.get('warning', '')}")
        sleep(5)


if __name__ == "__main__":
    main()
