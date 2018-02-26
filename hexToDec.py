#!/usr/bin/env python3
# -*- coding: utf-8 -*-

' Parser txt file tobe excel file.'

__author__ = 'Luan Yi Bo'

import xlwt
import argparse

PATH_TXT_FILE = "example.txt"

class HexToDec(object):
    un_valid_num = 208  # 16进制文档中无效的字节数
    ex_max_col_num = 16 # excel中列的个数

    # 是否为偶数
    def is_even(self, num):
        if num % 2 != 0:
            return False
        return True

    # 2个16进制合成10进制数据
    def synthesizer(self, num, data):
        return data[num] * 256 + data[num + 1]

    # 读取需要处理的数据
    def read_data(self, path=None):
        if path==None:
            return
        with open(path, "rb") as f:
            un_valid_buf = f.read(self.un_valid_num) # 预读处理
            need_parse_buf = f.read() # 需要处理
        return need_parse_buf

    # 处理数据
    def processing_data(self, need_parse_buf):
        decimal_num = []
        for num in range(len(need_parse_buf)):
            if not self.is_even(num):
                continue
            decimal_num.append(self.synthesizer(num, need_parse_buf))
        return decimal_num

    def parse(self, read_path=None, write_path=None):
        # 解析数据
        need_parse_buf = self.read_data(read_path)
        decimal_num    = self.processing_data(need_parse_buf)

        # 写成 excel
        self.write_xls_file(write_path, decimal_num)
        return

    def write_xls_file(self, path, data):
        # 创建工作簿
        f = xlwt.Workbook()
        # 创建sheet
        sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)

        # 写正文
        for row in range(int(len(data)/self.ex_max_col_num)+1):
            for col in range(self.ex_max_col_num):
                if row*16+col > len(data)-1:
                    break
                sheet1.write(row, col, data[row*16+col])

        # 保存文件
        if path == None:
            path = '16to10.xls'
        f.save(path)
        return


def main(args):
    read_path = args.read_file_txt
    write_path= args.write_excel

    ret = HexToDec()
    ret.parse(read_path, write_path)
    return

if __name__ == '__main__':
    arg_parser = argparse.ArgumentParser(description="解析16进制数据变成10进制。。。")
    arg_parser.add_argument("-read", "--read_file_txt", help="数据源")
    arg_parser.add_argument("-write" , "--write_excel", help="写入的文件名")
    args = arg_parser.parse_args()
    main(args)
