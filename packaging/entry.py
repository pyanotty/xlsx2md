"""PyInstaller / 単体実行用のエントリ。

パッケージの __main__.py は相対 import を使うため、スクリプトとして直接
凍結すると壊れる。ここでは絶対 import でエントリを呼ぶ。
"""
import sys

from xlsx2md.cli import main

if __name__ == "__main__":
    sys.exit(main())
