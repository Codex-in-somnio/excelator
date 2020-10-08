#!/usr/bin/env python3

import server
import argparse
import coloredlogs
import logging
import config
import os
import openpyxl
import re

SAVE_DIR_DEFAULT = 'replaced'

argparser = argparse.ArgumentParser()
argparser.add_argument('find', help='查找文本')
argparser.add_argument('replace', help='替换文本')
argparser.add_argument('-c', '--col',
                       help='列名（例如`E`），不填则针对全部可编辑列')
argparser.add_argument('-s', '--save-dir', default=SAVE_DIR_DEFAULT,
                       help=f'保存目录（默认：{SAVE_DIR_DEFAULT}）')
argparser.add_argument('-v', '--verbose', action='store_true',
                       help='打印调试日志')

args = argparser.parse_args()

HL_FMT = '\x1b[1;37;44m'
DIM_FMT = '\x1b[90m'
NORMAL_FMT = '\x1b[0m'

coloredlogs.install(level=logging.DEBUG if args.verbose else logging.INFO)


def highlight(text: str, pos_list: list, length: int):
    text_hl = ''
    cur_pos = 0
    for pos in sorted(pos_list):
        text_hl += text[cur_pos:pos]
        text_hl += HL_FMT
        text_hl += text[pos:pos+length]
        text_hl += NORMAL_FMT
        cur_pos = pos + length
    text_hl += text[cur_pos:]
    return text_hl


search_result = {}

for filename in os.listdir(server.DATA_PATH):
    if not filename.lower().endswith('.xlsx') or filename.startswith('~$'):
        logging.info(f'跳过文件：{filename}')
        continue
    logging.debug(f'打开：{filename}')
    search_result[filename] = server.search_xlsx_for(filename, args.find, True)

    if args.col:
        for location in set(search_result[filename].keys()):
            if location[0][0] != args.col.upper():
                del search_result[filename][location]

for filename in search_result:
    for location, hit in search_result[filename].items():
        sheet, cell_coord = location
        val, pos_list = hit
        print(DIM_FMT)
        print(f'文件名：{filename}，工作簿：{sheet}，单元格：{cell_coord}')
        print(NORMAL_FMT, end='')
        print(highlight(val, pos_list, len(args.find)))

answer = input('\n是否开始替换？(y/N)')
if not answer or answer.lower()[0] != 'y':
    exit()


find_pattern = re.compile(re.escape(args.find), re.IGNORECASE)
os.makedirs(args.save_dir, 0o755, True)

for filename in search_result:
    if not search_result[filename]:
        continue
    logging.debug(f'打开：{filename}')
    wb = openpyxl.open(os.path.join(server.DATA_PATH, filename))
    for location, hit in search_result[filename].items():
        sheet, cell_coord = location
        val, _ = hit
        wb[sheet][cell_coord].value = find_pattern.sub(args.replace, val)
    save_path = os.path.join(args.save_dir, filename)
    logging.debug(f'保存至路径：{save_path}')
    wb.save(save_path)
    wb.close()

logging.info('完成')
