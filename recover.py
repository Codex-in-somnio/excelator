#!/usr/bin/env python3

from server import db, DbCell, DATA_PATH

from datetime import datetime

import openpyxl
import argparse
import logging
import os
import json


def get_cells(filename, last_date=None):

    cells = db.session.query(
        DbCell.worksheet,
        DbCell.cell_coord,
        DbCell.text,
        db.func.max(DbCell.timestamp)
    ).filter(
        DbCell.filename == filename,
        not last_date or DbCell.timestamp < last_date
    ).group_by(
        DbCell.cell_coord,
        DbCell.worksheet,
        DbCell.filename
    ).all()

    worksheets = {}

    for worksheet_name, cell_coord, text, timestamp in cells:
        logging.info(f"最后编辑于: {timestamp.isoformat()}; "
                     f"工作簿: {worksheet_name}; "
                     f"坐标: {cell_coord}; "
                     f"文本: `{text}`")
        worksheets.setdefault(worksheet_name, [])
        worksheets[worksheet_name].append((cell_coord, text))

    return worksheets


def recover(source_xlsx_path, dest_xlsx_path, orig_xlsx_filename,
            last_date=None):
    worksheets = get_cells(orig_xlsx_filename, last_date)
    
    if not worksheets:
        logging.error("没有找到符合条件的数据")
        return

    logging.info(f"打开：{source_xlsx_path}")
    wb = openpyxl.open(source_xlsx_path)
    for worksheet_name, cells in worksheets.items():
        ws = wb[worksheet_name]
        for cell_coord, text in cells:
            ws[cell_coord] = text
    logging.info(f"保存至：{dest_xlsx_path}")
    wb.save(dest_xlsx_path)


if __name__ == "__main__":
    argparser = argparse.ArgumentParser()
    argparser.add_argument("source_xlsx_path",
                           help="源Excel路径（如空的翻译文件）")
    argparser.add_argument("dest_xlsx_path",
                           help="输出Excel路径")
    argparser.add_argument("-o", "--orig-xlsx-filename",
                           help="原data中的Excel文件名，"
                                "不提供此参数则用源Excel路径的文件名部分")
    argparser.add_argument("-d", "--date",
                           help="要恢复到的时间点，ISO8601格式，"
                                "不提供此参数则尽量恢复最新版本")
    args = argparser.parse_args()

    orig_filename = args.orig_xlsx_filename
    if not orig_filename:
        orig_filename = os.path.basename(args.source_xlsx_path)

    last_date = datetime.fromisoformat(args.date) if args.date else None

    recover(args.source_xlsx_path, args.dest_xlsx_path,
            orig_filename, last_date)
