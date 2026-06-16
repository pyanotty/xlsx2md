"""xlsx2md: Excel 文書(要件定義書・画面設計書)を Markdown に変換する。

要件定義書 v1.0 準拠。1 シート = 1 HTML テーブルとして忠実再現し、
シートごとに独立した .md を出力する。
"""

from .converter import Options, convert

__all__ = ["Options", "convert"]
__version__ = "1.0.0"
