import pandas as pd
import xlwings as xw
import matplotlib.pyplot as plt
import xlwt
from main_logic import Ui_Control
from DIDO_calibrate_test import DIDOCalibrate
from AI_calibrate_test import AICalibrate
from AO_calibrate_test import AOCalibrate
import time
#
class ExcelGenerate(Ui_Control,DIDOCalibrate,AICalibrate,AOCalibrate):
    def generateExcel(self, station, module):

        book = xlwt.Workbook(encoding='utf-8')
        sheet = book.add_sheet('校准校验表', cell_overwrite_ok=True)
        # 如果出现报错：Exception: Attempt to overwrite cell: sheetname='sheet1' rowx=0 colx=0
        # 需要加上：cell_overwrite_ok=True)
        # 这是因为重复操作一个单元格导致的
        sheet.col(0).width = 256 * 12
        for i in range(99):
            #     sheet.w
            #     tall_style = xlwt.easyxf('font:height 240;')  # 36pt,类型小初的字号
            first_row = sheet.row(i)
            first_row.height_mismatch = True
            first_row.height = 20 * 20

        # 为样式创建字体
        title_font = xlwt.Font()
        # 字体类型
        title_font.name = '宋'
        # 字体颜色
        title_font.colour_index = 0
        # 字体大小，11为字号，20为衡量单位
        title_font.height = 20 * 20
        # 字体加粗
        title_font.bold = True

        # 设置单元格对齐方式
        title_alignment = xlwt.Alignment()
        # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        title_alignment.horz = 0x02
        # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        title_alignment.vert = 0x01
        # 设置自动换行
        title_alignment.wrap = 1

        # 设置边框
        title_borders = xlwt.Borders()
        # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
        # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
        title_borders.left = 0
        title_borders.right = 0
        title_borders.top = 0
        title_borders.bottom = 0

        # # 设置背景颜色
        # pattern = xlwt.Pattern()
        # # 设置背景颜色的模式
        # pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        # # 背景颜色
        # pattern.pattern_fore_colour = i

        # # 初始化样式
        title_style = xlwt.XFStyle()
        title_style.borders = title_borders
        title_style.alignment = title_alignment
        title_style.font = title_font
        # # 设置文字模式
        # font.num_format_str = '#,##0.00'

        # sheet.write(i, 0, u'字体', style0)
        # sheet.write(i, 1, u'背景', style1)
        # sheet.write(i, 2, u'对齐方式', style2)
        # sheet.write(i, 3, u'边框', style3)

        # 合并单元格，合并第1行到第2行的第1列到第19列
        sheet.write_merge(0, 1, 0, 18, u'整机检验记录单V1.1', title_style)

        # row3
        # 为样式创建字体
        row3_font = xlwt.Font()
        # 字体类型
        row3_font.name = '宋'
        # 字体颜色
        row3_font.colour_index = 0
        # 字体大小，11为字号，20为衡量单位
        row3_font.height = 10 * 20
        # 字体加粗
        row3_font.bold = False

        # 设置单元格对齐方式
        row3_alignment = xlwt.Alignment()
        # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        row3_alignment.horz = 0x02
        # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        row3_alignment.vert = 0x01
        # 设置自动换行
        row3_alignment.wrap = 1

        # 设置边框
        row3_borders = xlwt.Borders()
        # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
        # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
        row3_borders.left = 0
        row3_borders.right = 0
        row3_borders.top = 0
        row3_borders.bottom = 0

        # # 设置背景颜色
        # pattern = xlwt.Pattern()
        # # 设置背景颜色的模式
        # pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        # # 背景颜色
        # pattern.pattern_fore_colour = i

        # # 初始化样式
        row3_style = xlwt.XFStyle()
        row3_style.borders = row3_borders
        row3_style.alignment = row3_alignment
        row3_style.font = row3_font
        # # 设置文字模式
        # font.num_format_str = '#,##0.00'

        # sheet.write(i, 0, u'字体', style0)
        # sheet.write(i, 1, u'背景', style1)
        # sheet.write(i, 2, u'对齐方式', style2)
        # sheet.write(i, 3, u'边框', style3)

        sheet.write_merge(2, 2, 0, 2, 'PN：', row3_style)
        sheet.write_merge(2, 2, 3, 5, f'{self.module_pn}', row3_style)
        sheet.write_merge(2, 2, 6, 8, 'SN：', row3_style)
        sheet.write_merge(2, 2, 9, 11, f'{self.module_sn}', row3_style)
        sheet.write_merge(2, 2, 12, 14, 'REV：', row3_style)
        sheet.write_merge(2, 2, 15, 17, f'{self.module_rev}', row3_style)

        # leftTitle
        # 为样式创建字体
        leftTitle_font = xlwt.Font()
        # 字体类型
        leftTitle_font.name = '宋'
        # 字体颜色
        leftTitle_font.colour_index = 0
        # 字体大小，11为字号，20为衡量单位
        leftTitle_font.height = 12 * 20
        # 字体加粗
        leftTitle_font.bold = True

        # 设置单元格对齐方式
        leftTitle_alignment = xlwt.Alignment()
        # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        leftTitle_alignment.horz = 0x02
        # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        leftTitle_alignment.vert = 0x01
        # 设置自动换行
        leftTitle_alignment.wrap = 1

        # 设置边框
        leftTitle_borders = xlwt.Borders()
        # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
        # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
        leftTitle_borders.left = 5
        leftTitle_borders.right = 5
        leftTitle_borders.top = 5
        leftTitle_borders.bottom = 5

        # # 设置背景颜色
        # pattern = xlwt.Pattern()
        # # 设置背景颜色的模式
        # pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        # # 背景颜色
        # pattern.pattern_fore_colour = i

        # # 初始化样式
        leftTitle_style = xlwt.XFStyle()
        leftTitle_style.borders = leftTitle_borders
        leftTitle_style.alignment = leftTitle_alignment
        leftTitle_style.font = leftTitle_font
        generalTest_row = 4
        sheet.write_merge(generalTest_row, generalTest_row + 1, 0, 0, '常规检测', leftTitle_style)
        CPU_row = 6
        sheet.write_merge(CPU_row, CPU_row + 7, 0, 0, 'CPU检测', leftTitle_style)
        DI_row = 14
        sheet.write_merge(DI_row, DI_row + 3, 0, 0, 'DI信号', leftTitle_style)
        DO_row = 18
        sheet.write_merge(DO_row, DO_row + 3, 0, 0, 'DO信号', leftTitle_style)
        AI_row = 22
        sheet.write_merge(AI_row, AI_row + 1, 0, 0, 'AI信号', leftTitle_style)
        AO_row = 24
        sheet.write_merge(AO_row, AO_row + 1, 0, 0, 'AO信号', leftTitle_style)
        result_row = 26
        sheet.write_merge(result_row, result_row + 1, 0, 3, '整体检测结果', leftTitle_style)

        # contentTitle
        # 为样式创建字体
        contentTitle_font = xlwt.Font()
        # 字体类型
        contentTitle_font.name = '宋'
        # 字体颜色
        contentTitle_font.colour_index = 0
        # 字体大小，11为字号，20为衡量单位
        contentTitle_font.height = 10 * 20
        # 字体加粗
        contentTitle_font.bold = False

        # 设置单元格对齐方式
        contentTitle_alignment = xlwt.Alignment()
        # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        contentTitle_alignment.horz = 0x02
        # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        contentTitle_alignment.vert = 0x01
        # 设置自动换行
        contentTitle_alignment.wrap = 1

        # 设置边框
        contentTitle_borders = xlwt.Borders()
        # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
        # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
        contentTitle_borders.left = 5
        contentTitle_borders.right = 5
        contentTitle_borders.top = 5
        contentTitle_borders.bottom = 5

        # # 设置背景颜色
        # pattern = xlwt.Pattern()
        # # 设置背景颜色的模式
        # pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        # # 背景颜色
        # pattern.pattern_fore_colour = i

        # # 初始化样式
        contentTitle_style = xlwt.XFStyle()
        contentTitle_style.borders = contentTitle_borders
        contentTitle_style.alignment = contentTitle_alignment
        contentTitle_style.font = contentTitle_font

        sheet.write_merge(generalTest_row, generalTest_row, 1, 2, '外观', contentTitle_style)
        sheet.write(generalTest_row, 3, '---', contentTitle_style)
        sheet.write_merge(generalTest_row, generalTest_row, 4, 5, 'Run指示灯', contentTitle_style)
        sheet.write(generalTest_row, 6, '---', contentTitle_style)
        sheet.write_merge(generalTest_row, generalTest_row, 7, 8, 'Error指示灯', contentTitle_style)
        sheet.write(generalTest_row, 9, '---', contentTitle_style)
        sheet.write_merge(generalTest_row, generalTest_row, 10, 11, 'CAN_Run指示灯', contentTitle_style)
        sheet.write(generalTest_row, 12, '---', contentTitle_style)
        sheet.write_merge(generalTest_row, generalTest_row, 13, 14, 'CAN_Error指示灯', contentTitle_style)
        sheet.write(generalTest_row, 15, '---', contentTitle_style)
        sheet.write_merge(generalTest_row, generalTest_row, 16, 17, '拨码（预留）', contentTitle_style)
        sheet.write(generalTest_row, 18, '---', contentTitle_style)
        sheet.write_merge(generalTest_row + 1, generalTest_row + 1, 1, 2, '非测试项', contentTitle_style)
        sheet.write(generalTest_row + 1, 3, '---', contentTitle_style)
        sheet.write_merge(generalTest_row + 1, generalTest_row + 1, 4, 5, '------', contentTitle_style)
        sheet.write(generalTest_row + 1, 6, '---', contentTitle_style)
        sheet.write_merge(generalTest_row + 1, generalTest_row + 1, 7, 8, '------', contentTitle_style)
        sheet.write(generalTest_row + 1, 9, '---', contentTitle_style)
        sheet.write_merge(generalTest_row + 1, generalTest_row + 1, 10, 11, '------', contentTitle_style)
        sheet.write(generalTest_row + 1, 12, '---', contentTitle_style)
        sheet.write_merge(generalTest_row + 1, generalTest_row + 1, 13, 14, '------', contentTitle_style)
        sheet.write(generalTest_row + 1, 15, '---', contentTitle_style)
        sheet.write_merge(generalTest_row + 1, generalTest_row + 1, 16, 17, '------', contentTitle_style)
        sheet.write(generalTest_row + 1, 18, '---', contentTitle_style)

        sheet.write_merge(CPU_row, CPU_row, 1, 2, '片外Flash读写', contentTitle_style)
        sheet.write(CPU_row, 3, '---', contentTitle_style)
        sheet.write_merge(CPU_row, CPU_row, 4, 5, 'MAC&序列号', contentTitle_style)
        sheet.write(CPU_row, CPU_row, '---', contentTitle_style)
        sheet.write_merge(CPU_row, CPU_row, 7, 8, '多功能按钮', contentTitle_style)
        sheet.write(CPU_row, 9, '---', contentTitle_style)
        sheet.write_merge(CPU_row, CPU_row, 10, 11, 'R/S拨杆', contentTitle_style)
        sheet.write(CPU_row, 12, '---', contentTitle_style)
        sheet.write_merge(CPU_row, CPU_row, 13, 14, '实时时钟', contentTitle_style)
        sheet.write(CPU_row, 15, '---', contentTitle_style)
        sheet.write_merge(CPU_row, CPU_row, 16, 17, 'SRAM', contentTitle_style)
        sheet.write(CPU_row, 18, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 1, CPU_row + 1, 1, 2, '掉电保存', contentTitle_style)
        sheet.write(CPU_row  + 1, 3, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 1, CPU_row + 1, 4, 5, 'U盘', contentTitle_style)
        sheet.write(CPU_row  + 1, 6, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 1, CPU_row + 1, 7, 8, 'type-C', contentTitle_style)
        sheet.write(CPU_row  + 1, 9, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 1, CPU_row + 1, 10, 11, 'RS232通讯', contentTitle_style)
        sheet.write(CPU_row  + 1, 12, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 1, CPU_row + 1, 13, 14, 'RS485通讯', contentTitle_style)
        sheet.write(CPU_row  + 1, 15, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 1, CPU_row + 1, 16, 17, 'CAN通讯(预留)', contentTitle_style)
        sheet.write(CPU_row  + 1, 18, '---', contentTitle_style)

        sheet.write_merge(CPU_row + 2, CPU_row + 4, 1, 2, '输入通道', contentTitle_style)
        sheet.write(CPU_row + 2, 3, '---', contentTitle_style)
        sheet.write(CPU_row + 2, 4, '---', contentTitle_style)
        sheet.write(CPU_row + 2, 5, '---', contentTitle_style)
        sheet.write(CPU_row + 2, 6, '---', contentTitle_style)
        sheet.write(CPU_row + 2, 7, '---', contentTitle_style)
        sheet.write(CPU_row + 2, 8, '---', contentTitle_style)
        sheet.write(CPU_row + 2, 9, '---', contentTitle_style)
        sheet.write(CPU_row + 2, 10, '---', contentTitle_style)
        sheet.write(CPU_row + 2, 11, '---', contentTitle_style)
        sheet.write(CPU_row + 2, 12, '---', contentTitle_style)
        sheet.write(CPU_row + 2, 13, '---', contentTitle_style)
        sheet.write(CPU_row + 2, 14, '---', contentTitle_style)
        sheet.write(CPU_row + 2, 15, '---', contentTitle_style)
        sheet.write(CPU_row + 2, 16, '---', contentTitle_style)
        sheet.write(CPU_row + 2, 17, '---', contentTitle_style)
        sheet.write(CPU_row + 2, 18, '---', contentTitle_style)

        sheet.write(CPU_row + 3, 3, '---', contentTitle_style)
        sheet.write(CPU_row + 3, 4, '---', contentTitle_style)
        sheet.write(CPU_row + 3, 5, '---', contentTitle_style)
        sheet.write(CPU_row + 3, 6, '---', contentTitle_style)
        sheet.write(CPU_row + 3, 7, '---', contentTitle_style)
        sheet.write(CPU_row + 3, 8, '---', contentTitle_style)
        sheet.write(CPU_row + 3, 9, '---', contentTitle_style)
        sheet.write(CPU_row + 3, 10, '---', contentTitle_style)
        sheet.write(CPU_row + 3, 11, '---', contentTitle_style)
        sheet.write(CPU_row + 3, 12, '---', contentTitle_style)
        sheet.write(CPU_row + 3, 13, '---', contentTitle_style)
        sheet.write(CPU_row + 3, 14, '---', contentTitle_style)
        sheet.write(CPU_row + 3, 15, '---', contentTitle_style)
        sheet.write(CPU_row + 3, 16, '---', contentTitle_style)
        sheet.write(CPU_row + 3, 17, '---', contentTitle_style)
        sheet.write(CPU_row + 3, 18, '---', contentTitle_style)

        sheet.write(CPU_row + 4, 3, '---', contentTitle_style)
        sheet.write(CPU_row + 4, 4, '---', contentTitle_style)
        sheet.write(CPU_row + 4, 5, '---', contentTitle_style)
        sheet.write(CPU_row + 4, 6, '---', contentTitle_style)
        sheet.write(CPU_row + 4, 7, '---', contentTitle_style)
        sheet.write(CPU_row + 4, 8, '---', contentTitle_style)
        sheet.write(CPU_row + 4, 9, '---', contentTitle_style)
        sheet.write(CPU_row + 4, 10, '---', contentTitle_style)
        sheet.write(CPU_row + 4, 11, '---', contentTitle_style)
        sheet.write(CPU_row + 4, 12, '---', contentTitle_style)
        sheet.write(CPU_row + 4, 13, '---', contentTitle_style)
        sheet.write(CPU_row + 4, 14, '---', contentTitle_style)
        sheet.write(CPU_row + 4, 15, '---', contentTitle_style)
        sheet.write(CPU_row + 4, 16, '---', contentTitle_style)
        sheet.write(CPU_row + 4, 17, '---', contentTitle_style)
        sheet.write(CPU_row + 4, 18, '---', contentTitle_style)

        sheet.write_merge(CPU_row + 5, CPU_row + 7, 1, 2, '输出通道', contentTitle_style)
        sheet.write(CPU_row + 5, 3, '---', contentTitle_style)
        sheet.write(CPU_row + 5, 4, '---', contentTitle_style)
        sheet.write(CPU_row + 5, 5, '---', contentTitle_style)
        sheet.write(CPU_row + 5, 6, '---', contentTitle_style)
        sheet.write(CPU_row + 5, 7, '---', contentTitle_style)
        sheet.write(CPU_row + 5, 8, '---', contentTitle_style)
        sheet.write(CPU_row + 5, 9, '---', contentTitle_style)
        sheet.write(CPU_row + 5, 10, '---', contentTitle_style)
        sheet.write(CPU_row + 5, 11, '---', contentTitle_style)
        sheet.write(CPU_row + 5, 12, '---', contentTitle_style)
        sheet.write(CPU_row + 5, 13, '---', contentTitle_style)
        sheet.write(CPU_row + 5, 14, '---', contentTitle_style)
        sheet.write(CPU_row + 5, 15, '---', contentTitle_style)
        sheet.write(CPU_row + 5, 16, '---', contentTitle_style)
        sheet.write(CPU_row + 5, 17, '---', contentTitle_style)
        sheet.write(CPU_row + 5, 18, '---', contentTitle_style)

        sheet.write(CPU_row + 6, 3, '---', contentTitle_style)
        sheet.write(CPU_row + 6, 4, '---', contentTitle_style)
        sheet.write(CPU_row + 6, 5, '---', contentTitle_style)
        sheet.write(CPU_row + 6, 6, '---', contentTitle_style)
        sheet.write(CPU_row + 6, 7, '---', contentTitle_style)
        sheet.write(CPU_row + 6, 8, '---', contentTitle_style)
        sheet.write(CPU_row + 6, 9, '---', contentTitle_style)
        sheet.write(CPU_row + 6, 10, '---', contentTitle_style)
        sheet.write(CPU_row + 6, 11, '---', contentTitle_style)
        sheet.write(CPU_row + 6, 12, '---', contentTitle_style)
        sheet.write(CPU_row + 6, 13, '---', contentTitle_style)
        sheet.write(CPU_row + 6, 14, '---', contentTitle_style)
        sheet.write(CPU_row + 6, 15, '---', contentTitle_style)
        sheet.write(CPU_row + 6, 16, '---', contentTitle_style)
        sheet.write(CPU_row + 6, 17, '---', contentTitle_style)
        sheet.write(CPU_row + 6, 18, '---', contentTitle_style)

        sheet.write(CPU_row + 7, 3, '---', contentTitle_style)
        sheet.write(CPU_row + 7, 4, '---', contentTitle_style)
        sheet.write(CPU_row + 7, 5, '---', contentTitle_style)
        sheet.write(CPU_row + 7, 6, '---', contentTitle_style)
        sheet.write(CPU_row + 7, 7, '---', contentTitle_style)
        sheet.write(CPU_row + 7, 8, '---', contentTitle_style)
        sheet.write(CPU_row + 7, 9, '---', contentTitle_style)
        sheet.write(CPU_row + 7, 10, '---', contentTitle_style)
        sheet.write(CPU_row + 7, 11, '---', contentTitle_style)
        sheet.write(CPU_row + 7, 12, '---', contentTitle_style)
        sheet.write(CPU_row + 7, 13, '---', contentTitle_style)
        sheet.write(CPU_row + 7, 14, '---', contentTitle_style)
        sheet.write(CPU_row + 7, 15, '---', contentTitle_style)
        sheet.write(CPU_row + 7, 16, '---', contentTitle_style)
        sheet.write(CPU_row + 7, 17, '---', contentTitle_style)
        sheet.write(CPU_row + 7, 18, '---', contentTitle_style)
        # DO
        sheet.write_merge(DO_row, DO_row, 1, 2, '通道号', contentTitle_style)
        sheet.write(DO_row, 3, 'CH1', contentTitle_style)
        sheet.write(DO_row, 4, 'CH2', contentTitle_style)
        sheet.write(DO_row, 5, 'CH3', contentTitle_style)
        sheet.write(DO_row, 6, 'CH4', contentTitle_style)
        sheet.write(DO_row, 7, 'CH5', contentTitle_style)
        sheet.write(DO_row, 8, 'CH6', contentTitle_style)
        sheet.write(DO_row, 9, 'CH7', contentTitle_style)
        sheet.write(DO_row, 10, 'CH8', contentTitle_style)
        sheet.write(DO_row, 11, 'CH9', contentTitle_style)
        sheet.write(DO_row, 12, 'CH10', contentTitle_style)
        sheet.write(DO_row, 13, 'CH11', contentTitle_style)
        sheet.write(DO_row, 14, 'CH12', contentTitle_style)
        sheet.write(DO_row, 15, 'CH13', contentTitle_style)
        sheet.write(DO_row, 16, 'CH14', contentTitle_style)
        sheet.write(DO_row, 17, 'CH15', contentTitle_style)
        sheet.write(DO_row, 18, 'CH16', contentTitle_style)
        sheet.write_merge(DO_row + 1, DO_row + 1, 1, 2, '是否合格', contentTitle_style)
        sheet.write(DO_row + 1, 3, '---', contentTitle_style)
        sheet.write(DO_row + 1, 4, '---', contentTitle_style)
        sheet.write(DO_row + 1, 5, '---', contentTitle_style)
        sheet.write(DO_row + 1, 6, '---', contentTitle_style)
        sheet.write(DO_row + 1, 7, '---', contentTitle_style)
        sheet.write(DO_row + 1, 8, '---', contentTitle_style)
        sheet.write(DO_row + 1, 9, '---', contentTitle_style)
        sheet.write(DO_row + 1, 10, '---', contentTitle_style)
        sheet.write(DO_row + 1, 11, '---', contentTitle_style)
        sheet.write(DO_row + 1, 12, '---', contentTitle_style)
        sheet.write(DO_row + 1, 13, '---', contentTitle_style)
        sheet.write(DO_row + 1, 14, '---', contentTitle_style)
        sheet.write(DO_row + 1, 15, '---', contentTitle_style)
        sheet.write(DO_row + 1, 16, '---', contentTitle_style)
        sheet.write(DO_row + 1, 17, '---', contentTitle_style)
        sheet.write(DO_row + 1, 18, '---', contentTitle_style)
        sheet.write_merge(DO_row + 2, DO_row + 2, 1, 2, '通道号', contentTitle_style)
        sheet.write(DO_row + 2, 3, 'CH17', contentTitle_style)
        sheet.write(DO_row + 2, 4, 'CH18', contentTitle_style)
        sheet.write(DO_row + 2, 5, 'CH19', contentTitle_style)
        sheet.write(DO_row + 2, 6, 'CH20', contentTitle_style)
        sheet.write(DO_row + 2, 7, 'CH21', contentTitle_style)
        sheet.write(DO_row + 2, 8, 'CH22', contentTitle_style)
        sheet.write(DO_row + 2, 9, 'CH23', contentTitle_style)
        sheet.write(DO_row + 2, 10, 'CH24', contentTitle_style)
        sheet.write(DO_row + 2, 11, 'CH25', contentTitle_style)
        sheet.write(DO_row + 2, 12, 'CH26', contentTitle_style)
        sheet.write(DO_row + 2, 13, 'CH27', contentTitle_style)
        sheet.write(DO_row + 2, 14, 'CH28', contentTitle_style)
        sheet.write(DO_row + 2, 15, 'CH29', contentTitle_style)
        sheet.write(DO_row + 2, 16, 'CH30', contentTitle_style)
        sheet.write(DO_row + 2, 17, 'CH31', contentTitle_style)
        sheet.write(DO_row + 2, 18, 'CH32', contentTitle_style)
        sheet.write_merge(DO_row + 3, DO_row + 3, 1, 2, '是否合格', contentTitle_style)
        sheet.write(DO_row + 3, 3, '---', contentTitle_style)
        sheet.write(DO_row + 3, 4, '---', contentTitle_style)
        sheet.write(DO_row + 3, 5, '---', contentTitle_style)
        sheet.write(DO_row + 3, 6, '---', contentTitle_style)
        sheet.write(DO_row + 3, 7, '---', contentTitle_style)
        sheet.write(DO_row + 3, 8, '---', contentTitle_style)
        sheet.write(DO_row + 3, 9, '---', contentTitle_style)
        sheet.write(DO_row + 3, 10, '---', contentTitle_style)
        sheet.write(DO_row + 3, 11, '---', contentTitle_style)
        sheet.write(DO_row + 3, 12, '---', contentTitle_style)
        sheet.write(DO_row + 3, 13, '---', contentTitle_style)
        sheet.write(DO_row + 3, 14, '---', contentTitle_style)
        sheet.write(DO_row + 3, 15, '---', contentTitle_style)
        sheet.write(DO_row + 3, 16, '---', contentTitle_style)
        sheet.write(DO_row + 3, 17, '---', contentTitle_style)
        sheet.write(DO_row + 3, 18, '---', contentTitle_style)

        # DI
        sheet.write_merge(DI_row, DI_row, 1, 2, '通道号', contentTitle_style)
        sheet.write(DI_row, 3, 'CH1', contentTitle_style)
        sheet.write(DI_row, 4, 'CH2', contentTitle_style)
        sheet.write(DI_row, 5, 'CH3', contentTitle_style)
        sheet.write(DI_row, 6, 'CH4', contentTitle_style)
        sheet.write(DI_row, 7, 'CH5', contentTitle_style)
        sheet.write(DI_row, 8, 'CH6', contentTitle_style)
        sheet.write(DI_row, 9, 'CH7', contentTitle_style)
        sheet.write(DI_row, 10, 'CH8', contentTitle_style)
        sheet.write(DI_row, 11, 'CH9', contentTitle_style)
        sheet.write(DI_row, 12, 'CH10', contentTitle_style)
        sheet.write(DI_row, 13, 'CH11', contentTitle_style)
        sheet.write(DI_row, 14, 'CH12', contentTitle_style)
        sheet.write(DI_row, 15, 'CH13', contentTitle_style)
        sheet.write(DI_row, 16, 'CH14', contentTitle_style)
        sheet.write(DI_row, 17, 'CH15', contentTitle_style)
        sheet.write(DI_row, 18, 'CH16', contentTitle_style)
        sheet.write_merge(DI_row + 1, DI_row + 1, 1, 2, '是否合格', contentTitle_style)
        sheet.write(DI_row + 1, 3, '---', contentTitle_style)
        sheet.write(DI_row + 1, 4, '---', contentTitle_style)
        sheet.write(DI_row + 1, 5, '---', contentTitle_style)
        sheet.write(DI_row + 1, 6, '---', contentTitle_style)
        sheet.write(DI_row + 1, 7, '---', contentTitle_style)
        sheet.write(DI_row + 1, 8, '---', contentTitle_style)
        sheet.write(DI_row + 1, 9, '---', contentTitle_style)
        sheet.write(DI_row + 1, 10, '---', contentTitle_style)
        sheet.write(DI_row + 1, 11, '---', contentTitle_style)
        sheet.write(DI_row + 1, 12, '---', contentTitle_style)
        sheet.write(DI_row + 1, 13, '---', contentTitle_style)
        sheet.write(DI_row + 1, 14, '---', contentTitle_style)
        sheet.write(DI_row + 1, 15, '---', contentTitle_style)
        sheet.write(DI_row + 1, 16, '---', contentTitle_style)
        sheet.write(DI_row + 1, 17, '---', contentTitle_style)
        sheet.write(DI_row + 1, 18, '---', contentTitle_style)
        sheet.write_merge(DI_row + 2, DI_row + 2, 1, 2, '通道号', contentTitle_style)
        sheet.write(DI_row + 2, 3, 'CH17', contentTitle_style)
        sheet.write(DI_row + 2, 4, 'CH18', contentTitle_style)
        sheet.write(DI_row + 2, 5, 'CH19', contentTitle_style)
        sheet.write(DI_row + 2, 6, 'CH20', contentTitle_style)
        sheet.write(DI_row + 2, 7, 'CH21', contentTitle_style)
        sheet.write(DI_row + 2, 8, 'CH22', contentTitle_style)
        sheet.write(DI_row + 2, 9, 'CH23', contentTitle_style)
        sheet.write(DI_row + 2, 10, 'CH24', contentTitle_style)
        sheet.write(DI_row + 2, 11, 'CH25', contentTitle_style)
        sheet.write(DI_row + 2, 12, 'CH26', contentTitle_style)
        sheet.write(DI_row + 2, 13, 'CH27', contentTitle_style)
        sheet.write(DI_row + 2, 14, 'CH28', contentTitle_style)
        sheet.write(DI_row + 2, 15, 'CH29', contentTitle_style)
        sheet.write(DI_row + 2, 16, 'CH30', contentTitle_style)
        sheet.write(DI_row + 2, 17, 'CH31', contentTitle_style)
        sheet.write(DI_row + 2, 18, 'CH32', contentTitle_style)
        sheet.write_merge(DI_row + 3, DI_row + 3, 1, 2, '是否合格', contentTitle_style)
        sheet.write(DI_row + 3, 3, '---', contentTitle_style)
        sheet.write(DI_row + 3, 4, '---', contentTitle_style)
        sheet.write(DI_row + 3, 5, '---', contentTitle_style)
        sheet.write(DI_row + 3, 6, '---', contentTitle_style)
        sheet.write(DI_row + 3, 7, '---', contentTitle_style)
        sheet.write(DI_row + 3, 8, '---', contentTitle_style)
        sheet.write(DI_row + 3, 9, '---', contentTitle_style)
        sheet.write(DI_row + 3, 10, '---', contentTitle_style)
        sheet.write(DI_row + 3, 11, '---', contentTitle_style)
        sheet.write(DI_row + 3, 12, '---', contentTitle_style)
        sheet.write(DI_row + 3, 13, '---', contentTitle_style)
        sheet.write(DI_row + 3, 14, '---', contentTitle_style)
        sheet.write(DI_row + 3, 15, '---', contentTitle_style)
        sheet.write(DI_row + 3, 16, '---', contentTitle_style)
        sheet.write(DI_row + 3, 17, '---', contentTitle_style)
        sheet.write(DI_row + 3, 18, '---', contentTitle_style)

        # AI
        sheet.write_merge(AI_row, AI_row + 1, 1, 1, '信号类型', contentTitle_style)
        sheet.write_merge(AI_row, AI_row + 1, 2, 2, '通道号', contentTitle_style)
        sheet.write_merge(AI_row, AI_row, 3, 5, '测试点1', contentTitle_style)
        sheet.write(AI_row + 1, 3, '理论值', contentTitle_style)
        sheet.write(AI_row + 1, 4, '测试值', contentTitle_style)
        sheet.write(AI_row + 1, 5, '精度', contentTitle_style)

        sheet.write_merge(AI_row, AI_row, 6, 8, '测试点2', contentTitle_style)
        sheet.write(AI_row + 1, 6, '理论值', contentTitle_style)
        sheet.write(AI_row + 1, 7, '测试值', contentTitle_style)
        sheet.write(AI_row + 1, 8, '精度', contentTitle_style)

        sheet.write_merge(AI_row, AI_row, 9, 11, '测试点3', contentTitle_style)
        sheet.write(AI_row + 1, 9, '理论值', contentTitle_style)
        sheet.write(AI_row + 1, 10, '测试值', contentTitle_style)
        sheet.write(AI_row + 1, 11, '精度', contentTitle_style)

        sheet.write_merge(AI_row, AI_row, 12, 14, '测试点4', contentTitle_style)
        sheet.write(AI_row + 1, 12, '理论值', contentTitle_style)
        sheet.write(AI_row + 1, 13, '测试值', contentTitle_style)
        sheet.write(AI_row + 1, 14, '精度', contentTitle_style)

        sheet.write_merge(AI_row, AI_row, 15, 17, '测试点5', contentTitle_style)

        sheet.write(AI_row + 1, 15, '理论值', contentTitle_style)
        sheet.write(AI_row + 1, 16, '测试值', contentTitle_style)
        sheet.write(AI_row + 1, 17, '精度', contentTitle_style)
        sheet.write(AI_row, 18, '', contentTitle_style)
        sheet.write(AI_row + 1, 18, '', contentTitle_style)

        # AO
        sheet.write_merge(AO_row, AO_row + 1, 1, 1, '信号类型', contentTitle_style)
        sheet.write_merge(AO_row, AO_row + 1, 2, 2, '通道号', contentTitle_style)
        sheet.write_merge(AO_row, AO_row, 3, 5, '测试点1', contentTitle_style)
        sheet.write(AO_row + 1, 3, '理论值', contentTitle_style)
        sheet.write(AO_row + 1, 4, '测试值', contentTitle_style)
        sheet.write(AO_row + 1, 5, '精度', contentTitle_style)

        sheet.write_merge(AO_row, AO_row, 6, 8, '测试点2', contentTitle_style)
        sheet.write(AO_row + 1, 6, '理论值', contentTitle_style)
        sheet.write(AO_row + 1, 7, '测试值', contentTitle_style)
        sheet.write(AO_row + 1, 8, '精度', contentTitle_style)

        sheet.write_merge(AO_row, AO_row, 9, 11, '测试点3', contentTitle_style)
        sheet.write(AO_row + 1, 9, '理论值', contentTitle_style)
        sheet.write(AO_row + 1, 10, '测试值', contentTitle_style)
        sheet.write(AO_row + 1, 11, '精度', contentTitle_style)

        sheet.write_merge(AO_row, AO_row, 12, 14, '测试点4', contentTitle_style)
        sheet.write(AO_row + 1, 12, '理论值', contentTitle_style)
        sheet.write(AO_row + 1, 13, '测试值', contentTitle_style)
        sheet.write(AO_row + 1, 14, '精度', contentTitle_style)

        sheet.write_merge(AO_row, AO_row, 15, 17, '测试点5', contentTitle_style)
        sheet.write(AO_row + 1, 15, '理论值', contentTitle_style)
        sheet.write(AO_row + 1, 16, '测试值', contentTitle_style)
        sheet.write(AO_row + 1, 17, '精度', contentTitle_style)
        sheet.write(AO_row, 18, '', contentTitle_style)
        sheet.write(AO_row + 1, 18, '', contentTitle_style)

        # 结果
        sheet.write_merge(result_row, result_row, 4, 5, '□ 合格', contentTitle_style)
        sheet.write_merge(result_row, result_row, 6, 18, ' ', contentTitle_style)
        sheet.write_merge(result_row + 1, result_row + 1, 4, 5, '□ 不合格', contentTitle_style)
        sheet.write_merge(result_row + 1, result_row + 1, 6, 18, ' ', contentTitle_style)

        # 补充说明
        sheet.write(result_row + 2, 0, '补充说明：', contentTitle_style)
        sheet.write_merge(result_row + 2, result_row + 2, 1, 18,
                          'AI/AO信号检验要记录数据，电压和电流的精度为2‰以下为合格、电阻的精度0.5℃以下合格；其他测试项合格打“√”，否则打“×”',
                          contentTitle_style)

        # 检测信息
        sheet.write_merge(result_row + 3, result_row + 3, 0, 1, '检验员：', contentTitle_style)
        sheet.write_merge(result_row + 3, result_row + 3, 2, 3, '555', contentTitle_style)
        sheet.write_merge(result_row + 3, result_row + 3, 4, 5, '检验日期：', contentTitle_style)
        sheet.write_merge(result_row + 3, result_row + 3, 6, 8, f'{time.strftime("%Y-%m-%d %H：%M：%S")}', contentTitle_style)
        sheet.write_merge(result_row + 3, result_row + 3, 9, 10, '审核：', contentTitle_style)
        sheet.write_merge(result_row + 3, result_row + 3, 11, 13, ' ', contentTitle_style)
        sheet.write_merge(result_row + 3, result_row + 3, 14, 15, '审核日期：', contentTitle_style)
        sheet.write_merge(result_row + 3, result_row + 3, 16, 18, ' ', contentTitle_style)
        if module == 'DI':
            self.fillInDIData(station, book, sheet)
        elif module == 'DO':
            self.fillInDOData(station, book, sheet)
        elif module == 'AI':
            self.fillInAIData(station, book, sheet)
        elif module == 'AO':
            pass

    # def fillInDIData(self,):
    #     sheet.write_merge(2, 2, 3, 5, '------', row3_style)