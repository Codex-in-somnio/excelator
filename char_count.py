#!/usr/bin/env python3

import openpyxl
import argparse
import os
import logging
import coloredlogs


argparser = argparse.ArgumentParser()
argparser.add_argument('dir', help='存有Excel的目录')
argparser.add_argument('col', help='列名（例如`E`）')
argparser.add_argument('-s', '--skip-rows', type=int, default=0,
                       help='跳过行数')
argparser.add_argument('-v', '--verbose', action='store_true',
                       help='打印调试日志')

args = argparser.parse_args()

coloredlogs.install(level=logging.DEBUG if args.verbose else logging.INFO)

col_index = ord(args.col.lower()) - ord('a')

total_char_count = 0

for filename in os.listdir(args.dir):
    if not filename.lower().endswith('.xlsx') or filename.startswith('~$'):
        logging.info(f'跳过文件：{filename}')
        continue
    logging.debug(f'打开：{filename}')
    wb = openpyxl.open(os.path.join(args.dir, filename))
    for ws in wb:
        logging.debug(f'工作簿：{ws.title}')
        rows = list(ws.iter_rows())
        if len(rows) <= args.skip_rows:
            continue
        for row in rows[args.skip_rows:]:
            text = row[col_index].value
            if text is None:
                continue
            cc = len(str(text))
            total_char_count += cc
            logging.debug(
                f'已统计文本: `{text}`，字数{cc}；当前总字数{total_char_count}')

logging.info(f'总字数：{total_char_count}')
