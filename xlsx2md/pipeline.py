"""構造復元パイプライン: ①抽出 → ②分割 → ③役割付与 → ⑤生成 → 検証。

「設計書として自然な Markdown」を出力する v2 のメイン経路。
"""

from __future__ import annotations

import logging
from pathlib import Path

from openpyxl import load_workbook

from .converter import Options, _extract_images, _sanitize
from .extract import extract_sheet
from .interpret import interpret
from .render import render_document
from .segment import segment
from .verify import verify_content

logger = logging.getLogger("xlsx2md")


def convert_structured(xlsx_path: str | Path, out_dir: str | Path,
                       options: Options | None = None) -> list[Path]:
    """xlsx を構造復元して自然な Markdown を出力し、生成パス一覧を返す。"""
    options = options or Options()
    xlsx_path = Path(xlsx_path)
    out_dir = Path(out_dir)

    wb_data = load_workbook(xlsx_path, data_only=True)
    wb_form = (load_workbook(xlsx_path, data_only=False)
               if options.formula_fallback else None)

    doc_dir = out_dir / _sanitize(xlsx_path.stem)
    images_dir = doc_dir / options.image_dirname
    doc_dir.mkdir(parents=True, exist_ok=True)

    produced: list[Path] = []
    for idx, name in enumerate(wb_data.sheetnames, start=1):
        ws = wb_data[name]
        wsf = wb_form[name] if wb_form is not None else None

        image_map = _extract_images(ws, images_dir, _sanitize(name), options)
        sheet = extract_sheet(ws, wsf, options)
        if sheet is None:
            md = f"# {name}\n\n_(空のシート)_\n"
        else:
            blocks = interpret(sheet, segment(sheet), image_map)
            md = render_document(sheet, blocks, options.image_dirname, image_map)
            _report_missing(sheet, md)

        out_path = doc_dir / f"{idx:02d}_{_sanitize(name)}.md"
        out_path.write_text(md, encoding="utf-8")
        produced.append(out_path)
        logger.info("wrote %s", out_path)

    return produced


def _report_missing(sheet, md: str) -> None:
    missing = verify_content(sheet, md)
    if missing:
        logger.warning("シート '%s': 出力に現れない元テキストが %d 件あります:",
                       sheet.name, len(missing))
        for m in missing:
            logger.warning("  - %s", m)
