from flask import Flask, render_template, request
from urllib.parse import quote_plus, unquote_plus
from flask_sqlalchemy import SQLAlchemy
from waitress import serve

import os
import openpyxl
import re
import threading
import json
import logging
import coloredlogs

import config


DATA_PATH = 'data'

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///data/backup/backup.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

lock = threading.Lock()

coloredlogs.install(level=logging.DEBUG)


class DbCell(db.Model):
    filename = db.Column(db.String, primary_key=True)
    worksheet = db.Column(db.String, primary_key=True)
    cell_coord = db.Column(db.String, primary_key=True)
    timestamp = db.Column(db.DateTime, primary_key=True,
                          server_default=db.func.now())
    text = db.Column(db.String, nullable=True)


class Cell:
    coord = None
    text = None
    merged = False
    skip = False
    edit = False
    style = ""
    tab_index = 0
    hidden = False

    def __init__(self, cell, merged_cell_ranges):
        self.c = cell
        for merged_cell_range in merged_cell_ranges:
            if cell.coordinate in merged_cell_range:
                if merged_cell_range.start_cell == cell:
                    self.merged = (
                        len(merged_cell_range.top),
                        len(merged_cell_range.left))
                else:
                    self.skip = True
                    return
        self.text = cell.value
        self.coord = cell.coordinate

        styles = {
            'overflow': 'hidden'
        }

        def get_color(color):
            if color.type == 'indexed':
                index = color.indexed - 1
                return '#' + openpyxl.styles.colors.COLOR_INDEX[index]
            elif color.type == 'rgb':
                return '#' + color.rgb
            return None

        # fg_color = get_color(cell.fill.fgColor)
        # if fg_color:
        #     styles['color'] = fg_color

        # bg_color = get_color(cell.fill.bgColor)
        # if bg_color:
        #     styles['background-color'] = bg_color

        for attr, value in styles.items():
            self.style += f'{attr}:{value};'

        if cell.column_letter in config.EDIT_COLS and \
                cell.row > config.HEADER_ROWS:
            self.edit = True
            self.tab_index = 100 + \
                cell.row + (cell.column - 1) * cell.parent.max_row

        if isinstance(cell.value, str) and cell.value[0] == '=':
            m = re.search(r'=LEN\(([A-Za-z0-9]+)\)', cell.value)
            if m:
                c = m.group(1)
                self.text = len(cell.parent[c].value)

        if cell.column_letter in config.HIDE_COLS:
            self.hidden = True


@app.template_filter('urlencode')
def urlencode_filter(s):
    return quote_plus(s)


@app.route('/')
def index():
    files = []
    for f in os.listdir(DATA_PATH):
        if os.path.isfile(os.path.join(DATA_PATH, f)) and \
                os.path.splitext(f)[1].lower() == '.xlsx':
            files.append(f)
    files.sort()
    cur_file = request.args.get('filename')
    cur_ws = request.args.get('worksheet')

    if not cur_file:
        return render_template('body.html', files=files)

    wb = openpyxl.open(os.path.join(DATA_PATH, cur_file))

    if cur_ws:
        ws = wb[cur_ws]
    else:
        ws = wb.active
        cur_ws = ws.title

    table = []
    msg = ''

    col_edge = 0
    row_edge = 0

    if ws.max_row > 1:
        for row in ws:
            _row = []
            row_empty = True
            for cell in row:
                _cell = Cell(cell, ws.merged_cells.ranges)
                _row.append(_cell)
                if cell.value or _cell.edit:
                    col_edge = max(col_edge, cell.column)
                    row_empty = False
            table.append(_row)
            if not row_empty:
                row_edge = cell.row

        table = table[:row_edge]
        for i in range(len(table)):
            table[i] = table[i][:col_edge]
    else:
        msg = '无内容'

    wss = [ws.title for ws in wb.worksheets]

    rendered = render_template('body.html',
                               files=files, cur_file=cur_file, cur_ws=cur_ws,
                               table=table, wss=wss, msg=msg)
    wb.close()
    return rendered


@app.route('/write', methods=['POST'])
def write():
    data = request.get_json()
    with lock:
        logging.debug(f'收到：{data}')
        wb = openpyxl.open(os.path.join(DATA_PATH, data['filename']))
        ws = wb[data['worksheet']]
        for cell, val in data['cells'].items():
            ws[cell].value = val
            db_cell = DbCell(filename=data['filename'],
                             worksheet=data['worksheet'],
                             cell_coord=cell, text=val)
            db.session.add(db_cell)
        db.session.commit()
        wb.save(os.path.join(DATA_PATH, data['filename']))
        wb.close()

    return 'ok'


if __name__ == '__main__':
    if not os.path.isfile('data/backup/backup.db'):
        logging.warning('备份数据库文件不存在，即将创建')
        db.create_all()
    serve(app, host='127.0.0.1', port=5000)
    # app.run(debug=True)
