#!/bin/env python3
"""
Copyright 2018 (C) Fabiano Engler Neto - fabianoengler(at)gmail(.)com

Simple script and class helper to execute a MySQL query and automatically
produce a nice formated XLSX output file.
"""

import sys
import mysql.connector
from openpyxl import Workbook
from openpyxl.utils.cell import get_column_letter
from openpyxl.worksheet.write_only import WriteOnlyCell
from openpyxl.styles import Font
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from decimal import Decimal
from math import ceil
from collections import Counter

verbosity_level = 2
#verbosity_level = 5

def verb(level, *args, **kwargs):
    kwargs['flush'] = True
    if level <= verbosity_level:
        print(*args, **kwargs)


class MySql2Xlsx(object):

    MIN_COL_WIDTH = 6
    MIN_COL_FULL_LEN = 25
    NUMBER_FORMAT = '#,##0.00'

    def __init__(self, mysql_config, query_str, out_fname, query_params=None):
        self.__mysql_config = mysql_config
        self.query_str = query_str 
        self.out_fname = out_fname
        self.query_params = query_params

    def generate_report(self):
        self.mysql_connect(self.__mysql_config)
        self.mysql_execute(self.query_str, self.query_params)
        self.create_workbook()
        self.prepare_sheet()
        self.fetch_rows_and_write()
        self.mysql_disconnect()
        self.make_final_adjustments()
        self.write_final_file()
        verb(1, 'Done.')

    def mysql_connect(self, mysql_config):
        verb(1, 'Connecting...')
        self.conn = mysql.connector.connect(**mysql_config)
        self.cursor = self.conn.cursor()
        return self.cursor

    def mysql_execute(self, query_str, query_params=None):
        verb(1, 'Executing SQL...')
        self.cursor.execute(query_str, query_params)
        if not len(self.cursor.column_names):
            raise ValueError('No columns on query set')
        return self.cursor

    def mysql_disconnect(self):
        self.conn.close()

    def create_workbook(self):
        verb(1, 'Creating XLSX Workbook...')
        self.wb = Workbook(write_only=True)
        return self.wb

    def create_worksheet(self):
        verb(3, 'Creating Worksheet...')
        self.ws = self.wb.create_sheet()
        return self.ws

    def freeze_panes(self):
        verb(3, 'Freezing first row...')
        self.ws.freeze_panes = 'A2'
        return self.ws

    def set_filters(self):
        self.ws.auto_filter.ref = 'A:{}'.format(
                get_column_letter(len(self.cursor.column_names)))
        return self.ws

    def prepare_sheet(self):
        self.create_worksheet()
        self.freeze_panes()
        self.set_filters()
        self.write_column_names()
        return self.ws

    def mysql_fetch_data_chunk(self, chunk_size=1000):
        return self.cursor.fetchmany(size=chunk_size)

    def mysql_fetch_rows_iterator(self, chunk_size=1000):
        while True:
            data_chunk = self.mysql_fetch_data_chunk(chunk_size)
            if not data_chunk:
                verb(2, '')
                return

            verb(2, '.', end='')
            for data_row in data_chunk:
                yield data_row

    def write_column_names(self):
        verb(3, 'Writing column names on first row...')
        sheet_row = []
        for col in self.cursor.column_names:
            col_name = col.decode('utf-8') if isinstance(col, bytes) else col
            val = col_name.replace('_', ' ').title()
            sheet_row.append(val)
        self.ws.append(sheet_row)
        return self.ws

    def fetch_rows_and_write(self, chunk_size=1000):
        cols_lengths = [ [] for _ in range(len(self.cursor.column_names)) ]
        cols_types = [ Counter() for _ in range(len(self.cursor.column_names)) ]

        verb(1, 'Writing data to worksheet...')
        for data_row in self.mysql_fetch_rows_iterator():
            sheet_row = []
            for i, value in enumerate(data_row):
                sheet_row.append(value)
                chars = 0 if value is None else len(str(value))
                if isinstance(value, int):
                    chars += 1
                elif isinstance(value, (float, Decimal)):
                    chars += 2
                cols_types[i][type(value)] += 1
                cols_lengths[i].append(0 if value is None else len(str(value)))
            self.ws.append(sheet_row)
        verb(2, 'Total Number of rows: {}'.format(self.cursor.rowcount))

        self.cols_lengths = cols_lengths 
        self.cols_types = cols_types 
        return self.ws

    def make_final_adjustments(self):
        verb(1, 'Final adjustments...')

        # need to save and re-open file to edit, as write_only mode does not
        # support in-memory editing
        verb(3, 'Saving intermediate file...')
        self.wb.save(self.out_fname)
        verb(3, 'Reloading intermediate file...')
        self.wb = load_workbook(self.out_fname)
        self.ws = self.wb.active

        self.resize_columns()
        self.format_numbers()
        self.format_column_names()

        return self.ws

    def resize_columns(self):
        verb(2, 'Resizing columns...')
        for i, col in enumerate(self.cols_lengths):
            m = max(col)
            if m <= self.MIN_COL_FULL_LEN:
                column_width = max(m, self.MIN_COL_WIDTH)
            else:
                dec9th = sorted(col)[ceil(len(col)/10*9)]
                column_width = min(m, dec9th)
            verb(5, 'column_width : {}'.format(column_width))
            
            column = get_column_letter(i+1)
            self.ws.column_dimensions[column].width = column_width

    def format_numbers(self):
        verb(2, 'Formating numbers...')
        for i, cells_types in enumerate(self.cols_types):
            del cells_types[type(None)]
            col_type = cells_types.most_common(1)[0][0]
            if col_type in (float, Decimal):
                for cell in self.ws[get_column_letter(i+1)]:
                    cell.number_format = self.NUMBER_FORMAT

    def format_column_names(self, font=None, alignment=None):
        verb(2, 'Formating column names...')
        alignment = Alignment(wrap_text=True) if alignment is None else alignment
        font = Font(bold=True) if font is None else font

        for i in range(len(self.cursor.column_names)):
            cell = self.ws.cell(row=1, column=i+1)
            cell.alignment = alignment 
            cell.font = font

        row1 = self.ws.row_dimensions[1]
        row1.height = 26

    def write_final_file(self):
        verb(1, 'Writing final file...')
        self.wb.save(self.out_fname)



def main():
    if not (2 <= len(sys.argv) <= 3) or sys.argv[1] in ('-h', '--help'):
        print('Usage:')
        print('    {} <query-file.sql> [output_file.xlsx]'.format(sys.argv[0]))
        sys.exit(-1)

    try:
        raw_sql = open(sys.argv[1], 'rt').read()
    except OSError as e:
        print(e)
        sys.exit(e.errno)

    from config import mysql_config 

    if len(sys.argv) == 3:
        out_fname = sys.argv[2]
    else:
        out_fname = '{}_result.xlsx'.format(sys.argv[1])

    mysql2xlsx = MySql2Xlsx(mysql_config, raw_sql, out_fname)
    mysql2xlsx.generate_report()



if __name__ == '__main__':
    main()






