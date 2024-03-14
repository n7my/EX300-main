# -*- coding: utf-8 -*-
import CAN_option
# import win32print
from ctypes import Array, c_ubyte
import time
import xlwt

# CAN帧结构体
# 发送帧ID
TRANSMIT_ID = 0x602

# 接收帧ID
RECEIVE_ID = 0x281

# 时间标识
TIME_STAMP = 0

# 是否使用时间标识
TIME_FLAG = 0

# 发送帧类型
TRANSMIT_SEND_TYPE = 1

# 接收帧类型
RECEIVE_SEND_TYPE = 1

# 是否是远程帧
REMOTE_FLAG = 0

# 是否是扩展帧
EXTERN_FLAG = 0

# 数据长度DLC
DATA_LEN = 8

# 用来接收的帧结构体数组的长度, 适配器中为每个通道设置了2000帧左右的接收缓存区
RECEIVE_LEN = 1

# 接收保留字段
WAIT_TIME = 500

# 要发送的帧结构体数组的长度(发送的帧数量), 最大为1000, 建议设为1, 每次发送单帧, 以提高发送效率
TRANSMIT_LEN = 1

ubyte_array_8 = c_ubyte * 8
DATA = ubyte_array_8(0, 0, 0, 0, 0, 0, 0, 0)
ubyte_array_3 = c_ubyte * 3
RESERVED_3 = ubyte_array_3(0, 0, 0)
m_can_obj = CAN_option.VCI_CAN_OBJ(RECEIVE_ID, TIME_STAMP, TIME_FLAG, RECEIVE_SEND_TYPE,
                                   REMOTE_FLAG, EXTERN_FLAG, DATA_LEN, DATA, RESERVED_3)


def isModulesOnline(CANAddr_array,module_array,waiting_time):
    # 检测设备心跳
    for i in range(len(module_array)):
        if not check_heartbeat(CANAddr_array[i], module_array[i], waiting_time):
            return False,i+1
    return True, 0
    #
    # if type == 'DI' or type == 'DO' or type == 'DIDO':
    #     if check_heartbeat(CANAddr_array[0], module_array[0], waiting_time) == False:
    #         return False,0
    # if type == 'AO' or type == 'AI' :
    #     if  check_heartbeat( CANAddr_relay, '继电器1',  waiting_time) == False:
    #         return False,1
    #     if  check_heartbeat( CANAddr_relay + 1, '继电器2', waiting_time) == False:
    #         return False,3
    # if  check_heartbeat( CAN2, module_2, waiting_time) == False:
    #     return False,7



def check_heartbeat(can_addr, inf, max_waiting):
    can_id = 0x700 + can_addr
    bool_receive,  m_can_obj = CAN_option.receiveCANbyID(can_id, max_waiting,1)
    # print( m_can_obj.Data)
    if bool_receive == False:
        # self.result_signal.emit(f'错误:未发现{inf}' + self.HORIZONTAL_LINE)
        return False
    # self.result_signal.emit(f'发现{inf}:收到心跳帧:{hex(self.m_can_obj.ID)}\n\n')
    return True

def generateExcel(code_array:list,station_array:list=[False,False,False],
                  channels:int=16,module:str=''):
    """

    :param self:
    :param code_array: 三码和MAC
    :param station_array: station_array[0]测试是否通过，station_array[1]是否进行电压测试，station_array[2]是否进行电流测试
    :param channels: 设备通道
    :param module: 设备类型:DI，DO，AI，AO，CPU
    :return:
    """
    try:
        # eID = 0
        book = xlwt.Workbook(encoding='utf-8')
        # eID = 1
        sheet = book.add_sheet('校准校验表', cell_overwrite_ok=True)
        # eID = 2
        # 如果出现报错:Exception: Attempt to overwrite cell: sheetname='sheet1' rowx=0 colx=0
        # 需要加上:cell_overwrite_ok=True)
        # 这是因为重复操作一个单元格导致的
        sheet.col(0).width = 256 * 12
        # eID = 3
        for i in range(99):
            # eID += 1
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
        # eID += 1
        # 字体加粗
        title_font.bold = True
        # eID += 1
        # 设置单元格对齐方式
        title_alignment = xlwt.Alignment()
        # eID += 1
        # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        title_alignment.horz = 0x02
        # eID += 1
        # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        title_alignment.vert = 0x01
        # eID += 1
        # 设置自动换行
        title_alignment.wrap = 1
        # eID += 1

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
        sheet.write_merge(2, 2, 0, 2, 'PN:', row3_style)
        sheet.write_merge(2, 2, 3, 5, f'{code_array[0]}', row3_style)
        sheet.write_merge(2, 2, 6, 8, 'SN:', row3_style)
        sheet.write_merge(2, 2, 9, 11, f'{code_array[1]}', row3_style)
        sheet.write_merge(2, 2, 12, 14, 'REV:', row3_style)
        sheet.write_merge(2, 2, 15, 17, f'{code_array[2]}', row3_style)

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
        sheet.write_merge(CPU_row, CPU_row + 8, 0, 0, 'CPU检测', leftTitle_style)
        DI_row = 15
        if module == 'DI':
            returnRow = DI_row
        sheet.write_merge(DI_row, DI_row + 3, 0, 0, 'DI信号', leftTitle_style)
        DO_row = 19
        if module == 'DO':
            returnRow = DO_row
        sheet.write_merge(DO_row, DO_row + 3, 0, 0, 'DO信号', leftTitle_style)
        AI_row = 23
        if module == 'AI':
            returnRow = AI_row
            if (station_array[1] and not station_array[2]):
                sheet.write_merge(AI_row, AI_row + 1 + channels*5, 0, 0, 'AI信号', leftTitle_style)
                AO_row = AI_row + 2 + channels*5
            elif not station_array[1] and station_array[2]:
                sheet.write_merge(AI_row, AI_row + 1 + channels * 2, 0, 0, 'AI信号', leftTitle_style)
                AO_row = AI_row + 2 + channels * 2
            elif station_array[1] and station_array[2]:
                sheet.write_merge(AI_row, AI_row + 1 + 7 * channels, 0, 0, 'AI信号', leftTitle_style)
                AO_row = AI_row + 2 + 7 * channels
            elif not station_array[1] and not station_array[2]:
                sheet.write_merge(AI_row, AI_row + 1, 0, 0, 'AI信号', leftTitle_style)
                AO_row = AI_row + 2
        else:
            sheet.write_merge(AI_row, AI_row + 1, 0, 0, 'AI信号', leftTitle_style)
            AO_row = AI_row + 2
        if module == 'AO':
            returnRow = AO_row
            if station_array[1] and not station_array[2]:
                sheet.write_merge(AO_row, AO_row + 1 + channels*5, 0, 0, 'AO信号', leftTitle_style)
                result_row = AO_row + 2 + channels*5
            elif not station_array[1] and station_array[2]:
                sheet.write_merge(AO_row, AO_row + 1 + channels * 2, 0, 0, 'AO信号', leftTitle_style)
                result_row = AO_row + 2 + channels * 2
            elif station_array[1] and station_array[2]:
                sheet.write_merge(AO_row, AO_row + 1 + 7 * channels, 0, 0, 'AO信号', leftTitle_style)
                result_row = AO_row + 2 + 7 * channels
            elif not station_array[1] and not station_array[2]:
                sheet.write_merge(AO_row, AO_row + 1, 0, 0, 'AO信号', leftTitle_style)
                result_row = AO_row + 2
        else:
            sheet.write_merge(AO_row, AO_row + 1, 0, 0, 'AO信号', leftTitle_style)
            result_row = AO_row + 2
        if module == 'CPU':
            returnRow = CPU_row

        sheet.write_merge(result_row, result_row + 1, 0, 3, '整机检测结果', leftTitle_style)

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

        sheet.write_merge(CPU_row, CPU_row, 1, 2, '型号', contentTitle_style)
        sheet.write(CPU_row, 3, '---', contentTitle_style)
        sheet.write_merge(CPU_row, CPU_row, 4, 5, 'SRAM', contentTitle_style)
        sheet.write(CPU_row, CPU_row, '---', contentTitle_style)
        sheet.write_merge(CPU_row, CPU_row, 7, 8, 'FLASH读写', contentTitle_style)
        sheet.write(CPU_row, 9, '---', contentTitle_style)
        sheet.write_merge(CPU_row, CPU_row, 10, 11, 'R/S拨杆', contentTitle_style)
        sheet.write(CPU_row, 12, '---', contentTitle_style)
        sheet.write_merge(CPU_row, CPU_row, 13, 14, 'MFK按键', contentTitle_style)
        sheet.write(CPU_row, 15, '---', contentTitle_style)
        sheet.write_merge(CPU_row, CPU_row, 16, 17, '掉电保存', contentTitle_style)
        sheet.write(CPU_row, 18, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 1, CPU_row + 1, 1, 2, 'RTC', contentTitle_style)
        sheet.write(CPU_row + 1, 3, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 1, CPU_row + 1, 4, 5, 'FPGA', contentTitle_style)
        sheet.write(CPU_row + 1, 6, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 1, CPU_row + 1, 7, 8, '各指示灯', contentTitle_style)
        sheet.write(CPU_row + 1, 9, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 1, CPU_row + 1, 10, 11, '输入通道', contentTitle_style)
        sheet.write(CPU_row + 1, 12, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 1, CPU_row + 1, 13, 14, '输出通道', contentTitle_style)
        sheet.write(CPU_row + 1, 15, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 1, CPU_row + 1, 16, 17, '以太网', contentTitle_style)
        sheet.write(CPU_row + 1, 18, '---', contentTitle_style)

        sheet.write_merge(CPU_row + 2, CPU_row + 2, 1, 2, 'RS-232C', contentTitle_style)
        sheet.write(CPU_row + 2, 3, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 2, CPU_row + 2, 4, 5, 'RS-485', contentTitle_style)
        sheet.write(CPU_row + 2, 6, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 2, CPU_row + 2, 7, 8, '右扩CAN', contentTitle_style)
        sheet.write(CPU_row + 2, 9, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 2, CPU_row + 2, 10, 11, 'MAC&三码', contentTitle_style)
        sheet.write(CPU_row + 2, 12, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 2, CPU_row + 2, 13, 14, '选项板', contentTitle_style)
        sheet.write(CPU_row + 2, 15, '---', contentTitle_style)
        sheet.write_merge(CPU_row + 2, CPU_row + 2, 16, 17, '------', contentTitle_style)
        sheet.write(CPU_row + 2, 18, '---', contentTitle_style)

        sheet.write_merge(CPU_row + 3, CPU_row + 6, 1, 2, '输入通道', contentTitle_style)
        sheet.write(CPU_row + 3, 3, '00', contentTitle_style)
        sheet.write(CPU_row + 3, 4, '01', contentTitle_style)
        sheet.write(CPU_row + 3, 5, '02', contentTitle_style)
        sheet.write(CPU_row + 3, 6, '03', contentTitle_style)
        sheet.write(CPU_row + 3, 7, '04', contentTitle_style)
        sheet.write(CPU_row + 3, 8, '05', contentTitle_style)
        sheet.write(CPU_row + 3, 9, '06', contentTitle_style)
        sheet.write(CPU_row + 3, 10, '07', contentTitle_style)
        sheet.write(CPU_row + 3, 11, '10', contentTitle_style)
        sheet.write(CPU_row + 3, 12, '11', contentTitle_style)
        sheet.write(CPU_row + 3, 13, '12', contentTitle_style)
        sheet.write(CPU_row + 3, 14, '13', contentTitle_style)
        sheet.write(CPU_row + 3, 15, '14', contentTitle_style)
        sheet.write(CPU_row + 3, 16, '15', contentTitle_style)
        sheet.write(CPU_row + 3, 17, '16', contentTitle_style)
        sheet.write(CPU_row + 3, 18, '17', contentTitle_style)
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

        sheet.write(CPU_row + 5, 3, '20', contentTitle_style)
        sheet.write(CPU_row + 5, 4, '21', contentTitle_style)
        sheet.write(CPU_row + 5, 5, '22', contentTitle_style)
        sheet.write(CPU_row + 5, 6, '23', contentTitle_style)
        sheet.write(CPU_row + 5, 7, '24', contentTitle_style)
        sheet.write(CPU_row + 5, 8, '25', contentTitle_style)
        sheet.write(CPU_row + 5, 9, '26', contentTitle_style)
        sheet.write(CPU_row + 5, 10, '27', contentTitle_style)
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

        sheet.write_merge(CPU_row + 7, CPU_row + 8, 1, 2, '输出通道', contentTitle_style)
        sheet.write(CPU_row + 7, 3, '00', contentTitle_style)
        sheet.write(CPU_row + 7, 4, '01', contentTitle_style)
        sheet.write(CPU_row + 7, 5, '02', contentTitle_style)
        sheet.write(CPU_row + 7, 6, '03', contentTitle_style)
        sheet.write(CPU_row + 7, 7, '04', contentTitle_style)
        sheet.write(CPU_row + 7, 8, '05', contentTitle_style)
        sheet.write(CPU_row + 7, 9, '06', contentTitle_style)
        sheet.write(CPU_row + 7, 10, '07', contentTitle_style)
        sheet.write(CPU_row + 7, 11, '10', contentTitle_style)
        sheet.write(CPU_row + 7, 12, '11', contentTitle_style)
        sheet.write(CPU_row + 7, 13, '12', contentTitle_style)
        sheet.write(CPU_row + 7, 14, '13', contentTitle_style)
        sheet.write(CPU_row + 7, 15, '14', contentTitle_style)
        sheet.write(CPU_row + 7, 16, '15', contentTitle_style)
        sheet.write(CPU_row + 7, 17, '16', contentTitle_style)
        sheet.write(CPU_row + 7, 18, '17', contentTitle_style)

        sheet.write(CPU_row + 8, 3, '---', contentTitle_style)
        sheet.write(CPU_row + 8, 4, '---', contentTitle_style)
        sheet.write(CPU_row + 8, 5, '---', contentTitle_style)
        sheet.write(CPU_row + 8, 6, '---', contentTitle_style)
        sheet.write(CPU_row + 8, 7, '---', contentTitle_style)
        sheet.write(CPU_row + 8, 8, '---', contentTitle_style)
        sheet.write(CPU_row + 8, 9, '---', contentTitle_style)
        sheet.write(CPU_row + 8, 10, '---', contentTitle_style)
        sheet.write(CPU_row + 8, 11, '---', contentTitle_style)
        sheet.write(CPU_row + 8, 12, '---', contentTitle_style)
        sheet.write(CPU_row + 8, 13, '---', contentTitle_style)
        sheet.write(CPU_row + 8, 14, '---', contentTitle_style)
        sheet.write(CPU_row + 8, 15, '---', contentTitle_style)
        sheet.write(CPU_row + 8, 16, '---', contentTitle_style)
        sheet.write(CPU_row + 8, 17, '---', contentTitle_style)
        sheet.write(CPU_row + 8, 18, '---', contentTitle_style)
        # DO
        sheet.write_merge(DO_row, DO_row, 1, 2, '通道号', contentTitle_style)
        sheet.write(DO_row, 3, '00', contentTitle_style)
        sheet.write(DO_row, 4, '01', contentTitle_style)
        sheet.write(DO_row, 5, '02', contentTitle_style)
        sheet.write(DO_row, 6, '03', contentTitle_style)
        sheet.write(DO_row, 7, '04', contentTitle_style)
        sheet.write(DO_row, 8, '05', contentTitle_style)
        sheet.write(DO_row, 9, '06', contentTitle_style)
        sheet.write(DO_row, 10, '07', contentTitle_style)
        sheet.write(DO_row, 11, '10', contentTitle_style)
        sheet.write(DO_row, 12, '11', contentTitle_style)
        sheet.write(DO_row, 13, '12', contentTitle_style)
        sheet.write(DO_row, 14, '13', contentTitle_style)
        sheet.write(DO_row, 15, '14', contentTitle_style)
        sheet.write(DO_row, 16, '15', contentTitle_style)
        sheet.write(DO_row, 17, '16', contentTitle_style)
        sheet.write(DO_row, 18, '17', contentTitle_style)
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
        sheet.write(DO_row + 2, 3, '20', contentTitle_style)
        sheet.write(DO_row + 2, 4, '21', contentTitle_style)
        sheet.write(DO_row + 2, 5, '22', contentTitle_style)
        sheet.write(DO_row + 2, 6, '23', contentTitle_style)
        sheet.write(DO_row + 2, 7, '24', contentTitle_style)
        sheet.write(DO_row + 2, 8, '25', contentTitle_style)
        sheet.write(DO_row + 2, 9, '26', contentTitle_style)
        sheet.write(DO_row + 2, 10, '27', contentTitle_style)
        sheet.write(DO_row + 2, 11, '30', contentTitle_style)
        sheet.write(DO_row + 2, 12, '31', contentTitle_style)
        sheet.write(DO_row + 2, 13, '32', contentTitle_style)
        sheet.write(DO_row + 2, 14, '33', contentTitle_style)
        sheet.write(DO_row + 2, 15, '34', contentTitle_style)
        sheet.write(DO_row + 2, 16, '35', contentTitle_style)
        sheet.write(DO_row + 2, 17, '36', contentTitle_style)
        sheet.write(DO_row + 2, 18, '37', contentTitle_style)
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
        sheet.write(DI_row, 3, '00', contentTitle_style)
        sheet.write(DI_row, 4, '01', contentTitle_style)
        sheet.write(DI_row, 5, '02', contentTitle_style)
        sheet.write(DI_row, 6, '03', contentTitle_style)
        sheet.write(DI_row, 7, '04', contentTitle_style)
        sheet.write(DI_row, 8, '05', contentTitle_style)
        sheet.write(DI_row, 9, '06', contentTitle_style)
        sheet.write(DI_row, 10, '07', contentTitle_style)
        sheet.write(DI_row, 11, '10', contentTitle_style)
        sheet.write(DI_row, 12, '11', contentTitle_style)
        sheet.write(DI_row, 13, 'C12', contentTitle_style)
        sheet.write(DI_row, 14, '13', contentTitle_style)
        sheet.write(DI_row, 15, '14', contentTitle_style)
        sheet.write(DI_row, 16, '15', contentTitle_style)
        sheet.write(DI_row, 17, '16', contentTitle_style)
        sheet.write(DI_row, 18, '17', contentTitle_style)
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
        sheet.write(DI_row + 2, 3, '20', contentTitle_style)
        sheet.write(DI_row + 2, 4, '21', contentTitle_style)
        sheet.write(DI_row + 2, 5, '22', contentTitle_style)
        sheet.write(DI_row + 2, 6, '23', contentTitle_style)
        sheet.write(DI_row + 2, 7, '24', contentTitle_style)
        sheet.write(DI_row + 2, 8, '25', contentTitle_style)
        sheet.write(DI_row + 2, 9, '26', contentTitle_style)
        sheet.write(DI_row + 2, 10, '27', contentTitle_style)
        sheet.write(DI_row + 2, 11, '30', contentTitle_style)
        sheet.write(DI_row + 2, 12, '31', contentTitle_style)
        sheet.write(DI_row + 2, 13, '32', contentTitle_style)
        sheet.write(DI_row + 2, 14, '33', contentTitle_style)
        sheet.write(DI_row + 2, 15, '34', contentTitle_style)
        sheet.write(DI_row + 2, 16, '35', contentTitle_style)
        sheet.write(DI_row + 2, 17, '36', contentTitle_style)
        sheet.write(DI_row + 2, 18, '37', contentTitle_style)
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
        sheet.write_merge(AI_row, AI_row + 1, 2, 2, '量程', contentTitle_style)
        sheet.write_merge(AI_row, AI_row + 1, 3, 3, '通道号', contentTitle_style)
        sheet.write_merge(AI_row, AI_row, 3 + 1, 5 + 1, '测试点1(100%)', contentTitle_style)
        sheet.write(AI_row + 1, 3 + 1, '理论值', contentTitle_style)
        sheet.write(AI_row + 1, 4 + 1, '测试值', contentTitle_style)
        sheet.write(AI_row + 1, 5 + 1, '精度', contentTitle_style)

        sheet.write_merge(AI_row, AI_row, 6 + 1, 8 + 1, '测试点2(75%)', contentTitle_style)
        sheet.write(AI_row + 1, 6 + 1, '理论值', contentTitle_style)
        sheet.write(AI_row + 1, 7 + 1, '测试值', contentTitle_style)
        sheet.write(AI_row + 1, 8 + 1, '精度', contentTitle_style)

        sheet.write_merge(AI_row, AI_row, 9 + 1, 11 + 1, '测试点3(50%)', contentTitle_style)
        sheet.write(AI_row + 1, 9 + 1, '理论值', contentTitle_style)
        sheet.write(AI_row + 1, 10 + 1, '测试值', contentTitle_style)
        sheet.write(AI_row + 1, 11 + 1, '精度', contentTitle_style)

        sheet.write_merge(AI_row, AI_row, 12 + 1, 14 + 1, '测试点4(25%)', contentTitle_style)
        sheet.write(AI_row + 1, 12 + 1, '理论值', contentTitle_style)
        sheet.write(AI_row + 1, 13 + 1, '测试值', contentTitle_style)
        sheet.write(AI_row + 1, 14 + 1, '精度', contentTitle_style)

        sheet.write_merge(AI_row, AI_row, 15 + 1, 17 + 1, '测试点5(0%)', contentTitle_style)

        sheet.write(AI_row + 1, 15 + 1, '理论值', contentTitle_style)
        sheet.write(AI_row + 1, 16 + 1, '测试值', contentTitle_style)
        sheet.write(AI_row + 1, 17 + 1, '精度', contentTitle_style)
        # sheet.write(AI_row, 18, '', contentTitle_style)
        # sheet.write(AI_row + 1, 18, '', contentTitle_style)

        # AO
        sheet.write_merge(AO_row, AO_row + 1, 1, 1, '信号类型', contentTitle_style)
        sheet.write_merge(AO_row, AO_row + 1, 2, 2, '量程', contentTitle_style)
        sheet.write_merge(AO_row, AO_row + 1, 3, 3, '通道号', contentTitle_style)
        sheet.write_merge(AO_row, AO_row, 3 + 1, 5 + 1, '测试点1(100%)', contentTitle_style)
        sheet.write(AO_row + 1, 3 + 1, '理论值', contentTitle_style)
        sheet.write(AO_row + 1, 4 + 1, '测试值', contentTitle_style)
        sheet.write(AO_row + 1, 5 + 1, '精度', contentTitle_style)

        sheet.write_merge(AO_row, AO_row, 6 + 1, 8 + 1, '测试点2(75%)', contentTitle_style)
        sheet.write(AO_row + 1, 6 + 1, '理论值', contentTitle_style)
        sheet.write(AO_row + 1, 7 + 1, '测试值', contentTitle_style)
        sheet.write(AO_row + 1, 8 + 1, '精度', contentTitle_style)

        sheet.write_merge(AO_row, AO_row, 9 + 1, 11 + 1, '测试点3(50%)', contentTitle_style)
        sheet.write(AO_row + 1, 9 + 1, '理论值', contentTitle_style)
        sheet.write(AO_row + 1, 10 + 1, '测试值', contentTitle_style)
        sheet.write(AO_row + 1, 11 + 1, '精度', contentTitle_style)

        sheet.write_merge(AO_row, AO_row, 12 + 1, 14 + 1, '测试点4(25%)', contentTitle_style)
        sheet.write(AO_row + 1, 12 + 1, '理论值', contentTitle_style)
        sheet.write(AO_row + 1, 13 + 1, '测试值', contentTitle_style)
        sheet.write(AO_row + 1, 14 + 1, '精度', contentTitle_style)

        sheet.write_merge(AO_row, AO_row, 15 + 1, 17 + 1, '测试点5(0%)', contentTitle_style)
        sheet.write(AO_row + 1, 15 + 1, '理论值', contentTitle_style)
        sheet.write(AO_row + 1, 16 + 1, '测试值', contentTitle_style)
        sheet.write(AO_row + 1, 17 + 1, '精度', contentTitle_style)
        # sheet.write(AO_row, 18, '', contentTitle_style)
        # sheet.write(AO_row + 1, 18, '', contentTitle_style)

        # 结果
        sheet.write_merge(result_row, result_row, 4, 5, '□ 合格', contentTitle_style)
        sheet.write_merge(result_row, result_row, 6, 18, ' ', contentTitle_style)
        sheet.write_merge(result_row + 1, result_row + 1, 4, 5, '□ 不合格', contentTitle_style)
        sheet.write_merge(result_row + 1, result_row + 1, 6, 18, ' ', contentTitle_style)

        # 补充说明
        sheet.write(result_row + 2, 0, '补充说明:', contentTitle_style)
        sheet.write_merge(result_row + 2, result_row + 2, 1, 18,
                          'AI/AO信号检验要记录数据，电压和电流的精度为1‰以下为合格；其他测试项合格打“√”，否则打“×”',
                          contentTitle_style)

        # 检测信息
        sheet.write_merge(result_row + 3, result_row + 3, 0, 1, '检验员:', contentTitle_style)
        sheet.write_merge(result_row + 3, result_row + 3, 2, 3, '555', contentTitle_style)
        sheet.write_merge(result_row + 3, result_row + 3, 4, 5, '检验日期:', contentTitle_style)
        sheet.write_merge(result_row + 3, result_row + 3, 6, 8, f'{time.strftime("%Y-%m-%d %H:%M:%S")}',
                          contentTitle_style)
        sheet.write_merge(result_row + 3, result_row + 3, 9, 10, '审核:', contentTitle_style)
        sheet.write_merge(result_row + 3, result_row + 3, 11, 13, ' ', contentTitle_style)
        sheet.write_merge(result_row + 3, result_row + 3, 14, 15, '审核日期:', contentTitle_style)
        sheet.write_merge(result_row + 3, result_row + 3, 16, 18, ' ', contentTitle_style)

        return True, book,sheet,returnRow
        # print("eID: ",eID)
    except Exception as e:
        print("Error:",e)
        return False, book,sheet,returnRow
        # print("eID: ", eID)

    # self.fillInAOData(station_array[0], book, sheet)
    # if module == 'DI':
    #     self.fillInDIData(station, book, sheet)
    # elif module == 'DO':
    #     self.fillInDOData(station, book, sheet)
    # elif module == 'AI':
    #     # # print('打印AI检测结果')
    #     self.fillInAIData(station, book, sheet)
    # elif module == 'AO':
    #     # # print('打印AI检测结果')
    #     self.fillInAOData(station, book, sheet)