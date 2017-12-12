# -*- coding: utf-8 -*-

# Created by David Teng on 17-12-4

import xlrd

# -*- coding: utf-8 -*-

from xlrd import open_workbook
from xlutils.copy import copy
import xlwt
import traceback


class excelOperations():
    u"""
        PYTHON 操作excel的一些操作
    """

    def __init__(self, excelFile):
        u"""
            提供错题列的第一个cell的坐标（startRow+1 就是填写评论的列）
            :param startRow:  起始行号
            :param startCol:  起始列号
        """
        self.filePath = excelFile
        self.rb = open_workbook(excelFile, formatting_info=True)  # excel文件句柄
        self.wb = copy(self.rb)             #
        self.wr = self.wb.get_sheet(0)  # excel文件中sheet句柄

        pass


    def insertBMP(self):
        
        try:
            self.wr.insert_bitmap("满分.bmp", 3, 6, 5, 5, 0.2, 0.1)
        except Exception as e:
            print e
            print traceback.format_exc()
        pass

    def setCellValue(self, rowN, colN, value, style=None):
        u"""
            设置指定cell的值 (对于合并的单元格，对其值的读取时，行号/列号第一个单元格)
            :param rowN:  行号
            :param colN:  列号
            :param value: 表格值
            :param wr: 写表格句柄
            :return: 返回True/False
        """
        try:

            self.wr.write(rowN, colN, value, style)
            return True
        except Exception as e:
            print e
            return False

        pass

    def getCellValue(self, rowN, colN):
        u"""
            获取指定cell的值，这里指错题的题号
            :param rowN:
            :param colN:
            :return: 返回cell的值
        """
        # print "获取数值"
        sh = self.rb.sheet_by_index(0)
        # print sh.cell(rowN, colN).ctype, "row:%s" % rowN
        tmpValue = sh.cell(rowN, colN).value
        if isinstance(tmpValue, float):
            tmpValue = str(int(tmpValue))
        return tmpValue
        pass


    def setFont(self, fontName=u"宋体", fontHeight=11, align=xlwt.Alignment.HORZ_LEFT, color="black"):
        """

            :param fontName: 字体，默认"宋体"
            :param fontHeight: 字号， 该参数的值等于字号*20
            :param align: 字符位置，具体
            :param color: 字符颜色
            :return:
        """
        style = xlwt.XFStyle()
        aligment = xlwt.Alignment()
        aligment.horz = align
        aligment.vert = xlwt.Alignment.VERT_CENTER
        aligment.wrap = xlwt.Alignment.WRAP_AT_RIGHT
        font = xlwt.Font()
        font.name = fontName
        font.colour_index = xlwt.Style.colour_map[color]
        font.height = fontHeight
        style.font = font
        style.alignment = aligment
        return style
        pass



    def saveFile(self):
        self.wb.save(self.filePath)

if __name__=="__main__":
    # main()
    pass

