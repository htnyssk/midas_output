"""配布用ランチャー。PyInstaller で exe 化するときのエントリポイント。

  midas_output.exe            メニューを表示
  midas_output.exe list       リスト(list.xlsx)と PNG を作成
  midas_output.exe word       list.xlsx から Word を作成
  midas_output.exe all        上記を続けて実行
"""
import sys

from make_list import generate_image_list
from main import create_professional_report

MENU = """
==============================================
  MIDAS 出力図 → Word 帳票作成ツール
==============================================
  1) 画像リスト作成   (images → PNG 変換 + list.xlsx)
  2) Word 作成        (list.xlsx → output.docx)
  3) 1 → 2 を続けて実行
  q) 終了
"""


def run(choice):
    if choice in ("1", "list", "3", "all"):
        generate_image_list()
    if choice in ("2", "word", "3", "all"):
        create_professional_report()


def main():
    if len(sys.argv) > 1:
        run(sys.argv[1].lower())
        return
    while True:
        print(MENU)
        try:
            choice = input("番号を入力してください: ").strip().lower()
        except EOFError:
            break
        if choice in ("q", "quit", "exit", ""):
            break
        if choice in ("1", "2", "3"):
            run(choice)
        else:
            print("1〜3 または q を入力してください。")


if __name__ == "__main__":
    main()
