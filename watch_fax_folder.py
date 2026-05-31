from pathlib import Path
from time import sleep

from app import incoming_dir, process_incoming_once


def main():
    folder = incoming_dir()
    folder.mkdir(exist_ok=True)
    print(f"受信FAXフォルダを監視中: {folder}")
    print("複合機のFAX保存先をこのフォルダにすると、確認待ち一覧へ自動で追加します。")
    while True:
        results = process_incoming_once()
        for result in results:
            if result["status"] == "pending_review":
                print(f"確認待ちに追加: {result['source']} -> {result['review_id']}")
            else:
                print(f"確認が必要: {result['source']} / {result.get('warning', '')}")
        sleep(5)


if __name__ == "__main__":
    main()
