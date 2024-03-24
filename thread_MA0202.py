# -*- coding: utf-8 -*-

import time

from PyQt5.QtCore import QThread, pyqtSignal

import CAN_option
from main_logic import *
from CAN_option import *
import otherOption
import logging
import struct

class MA0202Thread(QObject):
    result_signal = pyqtSignal(str)
    item_signal = pyqtSignal(list)
    pass_signal = pyqtSignal(bool)

    messageBox_signal = pyqtSignal(list)
    # excel_signal = pyqtSignal(list)
    allFinished_signal = pyqtSignal()
    label_signal = pyqtSignal(list)
    saveExcel_signal = pyqtSignal(list)
    print_signal = pyqtSignal(list)
    pic_messageBox_signal = pyqtSignal(list)
    moveToRow_signal = pyqtSignal(list)

    HORIZONTAL_LINE = "\n------------------------------------------------------------" \
                      "------------------------------------------------\n\n"

    appearance =False

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

    # 发送的数据
    ubyte_array_transmit = c_ubyte * 8
    m_transmitData = ubyte_array_transmit(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
    # 接收的数据
    ubyte_array_receive = c_ubyte * 8
    m_receiveData = ubyte_array_receive(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
    # serial_transmitData =[]
    # serial_receiveData = [0 for x in range(40)]
    # 主副线程弹窗结果
    result_queue = 0

    #各测试项通过标志
    isPassOp = True

    #是否取消后续所有测试
    isCancelAllTest = False



    def __init__(self, inf_MA0202list: list, result_queue):
        super().__init__()
        self.result_queue = result_queue
        self.is_running = True
        self.is_pause = False
        logging.basicConfig(filename='example.log', level=logging.DEBUG)
        # self.inf_MA0202list = [self.inf_param,self.inf_product, self.inf_CANIPAdrr,
        #                                 self.inf_serialPort, self.inf_MA0202_test]
        # self.inf_param
        self.MA0202_isChangePara = inf_MA0202list[0]
        self.moduleName_1 = inf_MA0202list[1]
        self.moduleName_2 = inf_MA0202list[2]
        self.MA0202_CANAddr1 = inf_MA0202list[3]
        self.MA0202_CANAddr2 = inf_MA0202list[4]
        self.isTestModule = inf_MA0202list[5]
        self.testModuleType = inf_MA0202list[6]
        self.serialPort_typeC = inf_MA0202list[7]
        self.current_dir = os.getcwd().replace('\\','/')+"/_internal"

        #信号等待时间(ms)
        self.waiting_time = 2000

        self.isCancelAllTest = False


        self.errorNum = 0
        self.errorInf = ''
        self.isPassAll = True
        self.pause_num = 1

        
        # #初始化表格状态
        # for i in range(len(self.MA0202_isTest)):
        #     self.result_signal.emit(f"self.MA0202_isTest[{i}]:{self.MA0202_isTest[i]}\n")
        #     print(f"self.MA0202_isTest[{i}]:{self.MA0202_isTest[i]}\n")
        #     if self.MA0202_isTest[i]:
        #         self.item_signal.emit([i, 3, 0, ''])
        #     else:
        #         self.item_signal.emit([i, 0, 0, ''])
        
    def MA0202(self):
        thread_signal = self.MA0202Option()
        if thread_signal:
            self.result_signal.emit('MA0202测试通过')
            self.label_signal.emit(['pass', '全部通过'])
        else:
            self.result_signal.emit('MA0202测试不通过')
            self.label_signal.emit(['fail', '未通过'])
        self.allFinished_signal.emit()
    def MA0202Option(self):
        try:
            # 写MA0202量程和码值
            w_MA0202_1 = [0xAC, 11, 0x00, 0x0F, 0x10, 0x00,
                          0x00, 0x00, 0x01, 0x00, 0x00,
                          0x00]
            # 写MA0202三码
            # w_MA0202_2 = [0xAC, 11, 0x00, 0x0F, 0x10, 0x00,
            #               0x00, 0x00, 0x01, 0x00, 0x00,
            #               0x00]
            # 读MA0202量程和码值
            r_MA0202_1 = [0xAC, 9, 0x00, 0x0F, 0x0E, 0x00, 0x00, 0x00, 0x00, 0x00]
            # 读MA0202三码
            r_MA0202_2 = [0xAC, 8, 0x00, 0x0F, 0x0E, 0x00, 0x00, 0xF1, 0x00]

            #   0-10V     0-20mA
            AI_range = [0x2A, 0x00, 0x34, 0x00]
            #   0-10V
            AO_range = [0xE8, 0x03]
            # 写AI量程0-10v
            w_MA0202_1[6] = 0x00
            w_MA0202_1[7] = 0x00
            w_MA0202_1[8] = 0x01
            w_MA0202_1[9] = 0x2A
            w_MA0202_1[10] = 0x00
            w_MA0202_1[11] = self.getCheckNum(w_MA0202_1[:11])
            self.pauseOption()
            if not self.is_running:
                self.isCancelAllTest = True
                return False
            # 1 写AI通道1量程0-10V
            self.CPU_optionPanel(transmitData=w_MA0202_1)
            # 2 写AI通道2量程0-10V
            w_MA0202_1[8] = 0x02
            w_MA0202_1[11] = self.getCheckNum(w_MA0202_1[:11])
            self.pauseOption()
            if not self.is_running:
                self.isCancelAllTest = True
                return False
            self.CPU_optionPanel(transmitData=w_MA0202_1)
            if not self.isCancelAllTest:
                self.pauseOption()
                if not self.is_running:
                    self.isCancelAllTest = True
                    return False
                self.result_signal.emit('------------写选项板AI通道1、2量程0-10V成功------------\n\n')
                # 3 读AI通道1量程0-10V
                r_MA0202_1[6] = 0x00
                r_MA0202_1[7] = 0x00
                r_MA0202_1[8] = 0x01
                r_MA0202_1[9] = self.getCheckNum(r_MA0202_1[:9])
                self.pauseOption()
                if not self.is_running:
                    self.isCancelAllTest = True
                    return False
                self.CPU_optionPanel(transmitData=r_MA0202_1, param1=[0x2A, 0x00])
                r_MA0202_1[8] = 0x02
                r_MA0202_1[9] = self.getCheckNum(r_MA0202_1[:9])
                self.pauseOption()
                if not self.is_running:
                    self.isCancelAllTest = True
                    return False
                self.CPU_optionPanel(transmitData=r_MA0202_1, param1=[0x2A, 0x00])
                if not self.isCancelAllTest:
                    self.pauseOption()
                    if not self.is_running:
                        self.isCancelAllTest = True
                        return False
                    self.result_signal.emit('------------读选项板AI通道1、2量程0-10V成功------------\n\n')
                    # 4 写AI量程通道1量程0-20mA
                    w_MA0202_1 = [0xAC, 11, 0x00, 0x0F, 0x10, 0x00, 0x00, 0x00, 0x01, 0x34, 0x00,
                                  0x00]
                    w_MA0202_1[11] = self.getCheckNum(w_MA0202_1[:11])
                    self.pauseOption()
                    if not self.is_running:
                        self.isCancelAllTest = True
                        return False
                    self.CPU_optionPanel(transmitData=w_MA0202_1)
                    # 5 写AI量程通道2量程0-20mA
                    w_MA0202_1 = [0xAC, 11, 0x00, 0x0F, 0x10, 0x00, 0x00, 0x00, 0x02, 0x34, 0x00,
                                  0x00]
                    w_MA0202_1[11] = self.getCheckNum(w_MA0202_1[:11])
                    self.pauseOption()
                    if not self.is_running:
                        self.isCancelAllTest = True
                        return False
                    self.CPU_optionPanel(transmitData=w_MA0202_1)
                    if not self.isCancelAllTest:
                        self.pauseOption()
                        if not self.is_running:
                            self.isCancelAllTest = True
                            return False
                        self.result_signal.emit('------------写选项板AI通道1、2量程0-20mA成功------------\n\n')
                        # 6 读AI通道1量程0-20mA
                        r_MA0202_1[6] = 0x00
                        r_MA0202_1[7] = 0x00
                        r_MA0202_1[8] = 0x01
                        r_MA0202_1[9] = self.getCheckNum(r_MA0202_1[:9])
                        self.pauseOption()
                        if not self.is_running:
                            self.isCancelAllTest = True
                            return False
                        self.CPU_optionPanel(transmitData=r_MA0202_1, param1=[0x34, 0x00])
                        # 7 读AI通道2量程0-20mA
                        r_MA0202_1[8] = 0x02
                        r_MA0202_1[9] = self.getCheckNum(r_MA0202_1[:9])
                        self.pauseOption()
                        if not self.is_running:
                            self.isCancelAllTest = True
                            return False
                        self.CPU_optionPanel(transmitData=r_MA0202_1, param1=[0x34, 0x00])
                        if not self.isCancelAllTest:
                            if not self.is_running:
                                self.isCancelAllTest = True
                                return False
                            self.result_signal.emit('------------读选项板AI通道1、2量程0-20mA成功------------\n\n')
                            try:
                                # 8控制AQ0004模块发送电流
                                CAN_bool, mess = self.CAN_init([1])
                            except Exception as e:
                                self.isPassOp &= False
                                if not self.is_running:
                                    self.isCancelAllTest = True
                                    return False
                                self.result_signal.emit(
                                    f'CAN_init error:{e}\nCAN设备通道1启动失败，后续测试全部取消。\n')
                                self.messageBox_signal.emit(['警告',
                                                             f'CAN_init error:{e}\nCAN设备通道1启动失败，'
                                                             f'后续测试全部取消。\n'])
                                reply = self.result_queue.get()
                                if reply == QMessageBox.AcceptRole:
                                    self.isCancelAllTest = True
                                    return False
                                else:
                                    self.isCancelAllTest = True
                                    return False
                                return False
                            if CAN_bool:
                                # 0v-10v的理论值
                                voltageTheory_0010 = [int(0x6c00 * 1.17), 0x6c00, (int)(0x6c00 * 0.75),
                                                      (int)(0x6c00 * 0.5), (int)(0x6c00 * 0.25),
                                                      0]  # 0x6c00 =>27648
                                arrayVol_0010 = ["11.7V测试", "10V测试", "7.5V测试", "5V测试",
                                                 "2.5V测试", "0V测试"]
                                # 0mA-20mA的理论值
                                currentTheory_0020 = [(int)(0x6c00 * 1.17), 0x6c00, (int)(0x6c00 * 0.75),
                                                      (int)(0x6c00 * 0.5), (int)(0x6c00 * 0.25),
                                                      0]  # 0x6c00 =>27648
                                arrayCur_0020 = ["23.4mA", "20mA", "15mA",
                                                 "10mA", "5mA", "0mA"]
                                AQ_transmit = (c_ubyte * 8)(0x00, 0x00, 0x00, 0x00,
                                                            0x00, 0x00, 0x00, 0x00)
                                for va in range(len(currentTheory_0020)):
                                    if va == 0:
                                        # 修改AQ量程为0-20mA
                                        AQ_transmit = [0x2b, 0x10, 0x63, 1,
                                                       0xeb, 0x03, 0x00, 0x00]

                                        bool_range1 = CAN_option.transmitCAN((0x600 + self.CANAddr5),
                                                                             AQ_transmit, 1)[0]
                                        AQ_transmit = [0x2b, 0x10, 0x63, 2,
                                                       0xeb, 0x03, 0x00, 0x00]
                                        bool_range2 = \
                                            CAN_option.transmitCAN((0x600 + self.CANAddr5),
                                                                   AQ_transmit, 1)[0]
                                        if bool_range1 & bool_range2:
                                            if not self.is_running:
                                                self.isCancelAllTest = True
                                                return False
                                            self.result_signal.emit(
                                                f"------------成功修改AQ0004通道1、2量程为0-20mA------------\n\n")
                                        else:
                                            if not self.is_running:
                                                self.isCancelAllTest = True
                                                return False
                                            self.result_signal.emit(
                                                f"------------修改AQ0004通道1、2量程为0-20mA失败，取消后续测试------------\n\n")
                                            self.isCancelAllTest = True
                                            self.isPassOp &= False
                                            return False
                                        time.sleep(0.5)
                                    for ch in range(2):
                                        # 发送0-20mA的电流
                                        AQ_transmit = [0x2b, 0x11, 0x64, ch + 1,
                                                       (currentTheory_0020[va] & 0xff),
                                                       ((currentTheory_0020[va] >> 8) & 0xff),
                                                       0x00, 0x00]
                                        for asd in range(10):
                                            CAN_option.transmitCAN((0x600 + self.CANAddr5),
                                                                   AQ_transmit, 1)
                                            # time.sleep(0.1)
                                        # 等待一秒，等待输出稳定
                                        self.result_signal.emit('\n等待输出稳定……\n')
                                        time.sleep(1)
                                        if not self.is_running:
                                            self.isCancelAllTest = True
                                            return False
                                        self.result_signal.emit(f'\n\n------------AQ0004通道{ch + 1}'
                                                                f'输出电流为：{arrayCur_0020[va]}'
                                                                f'------------\n')
                                        r_MA0202_1 = [0xAC, 9, 0x00, 0x0F, 0x0E, 0x00, 0x00,
                                                      0x01, ch + 1, 0x00]
                                        r_MA0202_1[9] = self.getCheckNum(r_MA0202_1[:9])
                                        rate = round(float(27648 / 4095), 2)
                                        self.pauseOption()
                                        if not self.is_running:
                                            self.isCancelAllTest = True
                                            return False
                                        self.CPU_optionPanel(transmitData=r_MA0202_1,
                                                             param1=[(int(currentTheory_0020[va] / rate) & 0xff),
                                                                     ((int(currentTheory_0020[
                                                                               va] / rate) >> 8) & 0xff)])
                                    if self.isCancelAllTest:
                                        return False
                            else:
                                self.isPassOp &= False
                                self.messageBox_signal.emit(mess)
                                reply = self.result_queue.get()
                                if reply == QMessageBox.AcceptRole:
                                    self.isCancelAllTest = True
                                    return False
                                else:
                                    self.isCancelAllTest = True
                                    return False
                        else:
                            return False
                    else:
                        return False
                else:
                    return False
            else:
                return False
            # AO
            if not self.isCancelAllTest:
                # 写AO量程0-10v
                w_MA0202_1[6] = 0x01
                w_MA0202_1[7] = 0x00
                w_MA0202_1[8] = 0x01
                w_MA0202_1[9] = 0xE8
                w_MA0202_1[10] = 0x03
                w_MA0202_1[11] = self.getCheckNum(w_MA0202_1[:11])
                self.pauseOption()
                if not self.is_running:
                    self.isCancelAllTest = True
                    return False
                # 1 写AO通道1量程0-10V
                self.CPU_optionPanel(transmitData=w_MA0202_1)
                # 2 写AI通道2量程0-10V
                w_MA0202_1[8] = 0x02
                w_MA0202_1[11] = self.getCheckNum(w_MA0202_1[:11])
                self.pauseOption()
                if not self.is_running:
                    self.isCancelAllTest = True
                    return False
                self.CPU_optionPanel(transmitData=w_MA0202_1)
                if not self.isCancelAllTest:
                    if not self.is_running:
                        self.isCancelAllTest = True
                        return False
                    self.result_signal.emit(
                        '------------写选项板AO通道1、2量程0-10V成功------------\n\n')
                    # 3 读AO通道1量程0-10V
                    r_MA0202_1[6] = 0x01
                    r_MA0202_1[7] = 0x00
                    r_MA0202_1[8] = 0x01
                    r_MA0202_1[9] = self.getCheckNum(r_MA0202_1[:9])
                    self.pauseOption()
                    if not self.is_running:
                        self.isCancelAllTest = True
                        return False
                    self.CPU_optionPanel(transmitData=r_MA0202_1, param1=[0xE8, 0x03])
                    r_MA0202_1[8] = 0x02
                    r_MA0202_1[9] = self.getCheckNum(r_MA0202_1[:9])
                    self.pauseOption()
                    if not self.is_running:
                        self.isCancelAllTest = True
                        return False
                    self.CPU_optionPanel(transmitData=r_MA0202_1, param1=[0xE8, 0x03])
                    if not self.isCancelAllTest:
                        self.pauseOption()
                        if not self.is_running:
                            self.isCancelAllTest = True
                            return False
                        self.result_signal.emit(
                            '------------读选项板AO通道1、2量程0-10V成功------------\n\n')
                        try:
                            # 8控制选项板AO通道发送电压
                            CAN_bool, mess = self.CAN_init([1])
                            if CAN_bool:
                                # 0v-10v的理论值
                                voltageTheory_0010 = [int(4095 * 1.17), 4095,
                                                      (int)(4095 * 0.75),
                                                      (int)(4095 * 0.5), (int)(4095 * 0.25),
                                                      0]  # 0x6c00 =>27648
                                # voltageTheory_0010 = [0]
                                arrayVol_0010 = ["11.7V测试", "10V测试", "7.5V测试", "5V测试",
                                                 "2.5V测试", "0V测试"]

                                AE_transmit = (c_ubyte * 8)(0x00, 0x00, 0x00, 0x00,
                                                            0x00, 0x00, 0x00, 0x00)
                                for va in range(len(voltageTheory_0010)):
                                    if va == 0:
                                        # 修改AI量程为0-10V
                                        AE_transmit = [0x2b, 0x10, 0x61, 1,
                                                       0x2a, 0x00, 0x00, 0x00]
                                        bool_range1 = \
                                            CAN_option.transmitCAN((0x600 + self.CANAddr5),
                                                                   AE_transmit, 1)[0]
                                        AE_transmit = [0x2b, 0x10, 0x61, 2,
                                                       0x2a, 0x00, 0x00, 0x00]
                                        bool_range2 = \
                                            CAN_option.transmitCAN((0x600 + self.CANAddr5),
                                                                   AE_transmit, 1)[0]
                                        if bool_range1 & bool_range2:
                                            if not self.is_running:
                                                self.isCancelAllTest = True
                                                return False
                                            self.result_signal.emit(
                                                f"------------成功修改AE0400通道1、2量程为0-10V------------\n\n")
                                        else:
                                            if not self.is_running:
                                                self.isCancelAllTest = True
                                                return False
                                            self.result_signal.emit(
                                                f"------------修改AE0400通道1、2量程为0-10V失败，取消后续测试------------\n\n")
                                            self.isCancelAllTest = True
                                            self.isPassOp &= False
                                            return False

                                    for ch in range(2):
                                        # 发送0-10V的电压
                                        w_MA0202_1 = [0xAC, 11, 0x00, 0x0F, 0x10, 0x00,
                                                      0x01, 0x01, ch + 1,
                                                      (voltageTheory_0010[va] & 0xff),
                                                      ((voltageTheory_0010[va] >> 8) & 0xff), 0x00]
                                        if va == 5:
                                            w_MA0202_1[9] = 0x00
                                            w_MA0202_1[10] = 0x00
                                        w_MA0202_1[11] = self.getCheckNum(w_MA0202_1[:11])
                                        self.pauseOption()
                                        if not self.is_running:
                                            self.isCancelAllTest = True
                                            return False
                                        self.CPU_optionPanel(transmitData=w_MA0202_1)
                                        if not self.isCancelAllTest:
                                            if not self.is_running:
                                                self.isCancelAllTest = True
                                                return False
                                            self.result_signal.emit(
                                                f'\n\n------------选项板AO通道{ch + 1}'
                                                f'输出电压为：{arrayVol_0010[va]}'
                                                f'------------')
                                            if not self.is_running:
                                                self.isCancelAllTest = True
                                                return False
                                            self.result_signal.emit('等待接收的信号稳定……\n\n')
                                            tt0 = time.time()
                                            while True:
                                                rate = round(float(27648 / 4095), 2)
                                                if (time.time() - tt0) * 1000 > 10000:
                                                    if not self.is_running:
                                                        self.isCancelAllTest = True
                                                        return False
                                                    self.result_signal.emit(
                                                        '\n\n*********AE0400无法接收到稳定数据，'
                                                        '取消后续测试！*********\n\n')
                                                    self.isPassOp &= False
                                                    # self.isCancelAllTest = True
                                                    return False
                                                bool_AE, AE_value = self.receiveAIData(CANAddr_AI=self.CANAddr4,
                                                                                       channel=ch)

                                                if abs(voltageTheory_0010[va] - int(AE_value / rate)) < 40:
                                                    return False
                                            if self.isCancelAllTest:
                                                return False

                                            tt1 = time.time()
                                            while True:
                                                if (time.time() - tt1) * 1000 > 2000:
                                                    self.isPassOp &= False
                                                    if not self.is_running:
                                                        self.isCancelAllTest = True
                                                        return False
                                                    self.result_signal.emit(
                                                        '\n\n*********AE0400未接收到数据，取消后续测试！*********\n\n')
                                                    # self.isCancelAllTest = True
                                                    return False
                                                try:
                                                    bool_AE, AE_value = self.receiveAIData(CANAddr_AI=self.CANAddr4,
                                                                                           channel=ch)
                                                except:
                                                    if not self.is_running:
                                                        self.isCancelAllTest = True
                                                        return False
                                                    self.result_signal.emit(
                                                        f'ErrorInf:\n{traceback.format_exc()}')
                                                if bool_AE:
                                                    if not self.is_running:
                                                        self.isCancelAllTest = True
                                                        return False
                                                    self.result_signal.emit(
                                                        f'-----------模拟量输出通道{ch + 1}写入的值为{voltageTheory_0010[va]}-----------。\n')
                                                    if not self.is_running:
                                                        self.isCancelAllTest = True
                                                        return False
                                                    self.result_signal.emit(
                                                        f'-----------AE0400接收的值为{int(AE_value)}。-----------\n')
                                                    self.result_signal.emit(
                                                        f'-----------AE0400接收值转换比例后的值为{int(AE_value / rate)}。-----------\n')
                                                    try:
                                                        if abs(voltageTheory_0010[va] - int(AE_value / rate)) < 40:
                                                            self.isPassOp &= True
                                                            return False
                                                        else:
                                                            self.isPassOp &= False
                                                            return False
                                                    except:
                                                        if not self.is_running:
                                                            self.isCancelAllTest = True
                                                            return False

                                                        self.result_signal.emit(
                                                            f'ErrorInf:\n{traceback.format_exc()}')
                                                        return False
                                                else:
                                                    continue
                                    if self.isCancelAllTest:
                                        return False
                            else:
                                self.isPassOp &= False
                                self.messageBox_signal.emit(mess)
                                reply = self.result_queue.get()
                                if reply == QMessageBox.AcceptRole:
                                    self.isCancelAllTest = True
                                    return False
                                else:
                                    self.isCancelAllTest = True
                                    return False
                                return False
                        except Exception as e:
                            self.isPassOp &= False
                            if not self.is_running:
                                self.isCancelAllTest = True
                                return False

                            self.result_signal.emit(
                                f'CAN_init error:{e}\nCAN设备通道1启动失败，后续测试全部取消。\n')
                            self.messageBox_signal.emit(['警告',
                                                         f'CAN_init error:{e}\nCAN设备通道1启动失败，'
                                                         f'后续测试全部取消。\n'])

                            reply = self.result_queue.get()
                            if reply == QMessageBox.AcceptRole:
                                self.isCancelAllTest = True
                                return False
                            else:
                                self.isCancelAllTest = True
                                return False
                            return False
            else:
                return False

            # 三码
            if not self.isCancelAllTest:
                option_pn = 'P0000010390631'
                option_sn = 'S1223-001083'
                option_rev = '06'
                w_PN = [0xac, 0x21, 0x00, 0x0f, 0x10, 0x00, 0x00, 0xf1, 0x18,
                        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                        0x00]
                for pn in range(len(option_pn)):
                    w_PN[9 + pn] = ord(option_pn[pn])
                w_PN[33] = self.getCheckNum(w_PN[:33])
                if not self.is_running:
                    self.isCancelAllTest = True
                    return False

                self.result_signal.emit(f'开始写入选项板PN码：{option_pn}。\n')
                self.pauseOption()
                if not self.is_running:
                    self.isCancelAllTest = True
                    return False
                self.CPU_optionPanel(transmitData=w_PN)
                if not self.isCancelAllTest:
                    if not self.is_running:
                        self.isCancelAllTest = True
                        return False

                    self.result_signal.emit('写入PN成功！\n')
                    self.result_signal.emit('开始读取选项板PN码。\n\n')
                    r_PN = [0xac, 8, 0x00, 0x0f, 0x0e, 0x00, 0x00, 0xf1, 0x00]
                    r_PN[8] = self.getCheckNum(r_PN[:8])
                    self.pauseOption()
                    if not self.is_running:
                        self.isCancelAllTest = True
                        return False
                    self.CPU_optionPanel(transmitData=r_PN, param2=option_pn)
                    if not self.isCancelAllTest:
                        if not self.is_running:
                            self.isCancelAllTest = True
                            return False

                        self.result_signal.emit('读PN成功！\n\n')
                        w_SN = [0xac, 0x19, 0x00, 0x0f, 0x10, 0x00, 0x00, 0xf2, 0x10,
                                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                0x00]
                        for sn in range(len(option_sn)):
                            w_SN[9 + sn] = ord(option_sn[sn])
                        w_SN[25] = self.getCheckNum(w_SN[:25])
                        # self.result_signal.emit(f'w_SN:{w_SN}')
                        if not self.is_running:
                            self.isCancelAllTest = True
                            return False

                        self.result_signal.emit(f'开始写入选项板SN码：{option_sn}。\n')
                        self.pauseOption()
                        if not self.is_running:
                            self.isCancelAllTest = True
                            return False
                        self.CPU_optionPanel(transmitData=w_SN)
                        if not self.isCancelAllTest:
                            if not self.is_running:
                                self.isCancelAllTest = True
                                return False

                            self.result_signal.emit('写入SN成功！\n')
                            self.result_signal.emit('开始读取选项板SN码。\n\n')
                            r_SN = [0xac, 8, 0x00, 0x0f, 0x0e, 0x00, 0x00, 0xf2, 0x00]
                            r_SN[8] = self.getCheckNum(r_SN[:8])
                            self.pauseOption()
                            if not self.is_running:
                                self.isCancelAllTest = True
                                return False
                            self.CPU_optionPanel(transmitData=r_SN, param2=option_sn)
                            if not self.isCancelAllTest:
                                if not self.is_running:
                                    self.isCancelAllTest = True
                                    return False

                                self.result_signal.emit('读SN成功！\n\n')
                                w_REV = [0xac, 0x11, 0x00, 0x0f, 0x10, 0x00, 0x00, 0xf3, 0x08,
                                         0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                         0x00]
                                for rev in range(len(option_rev)):
                                    w_REV[9 + rev] = ord(option_rev[rev])
                                w_REV[17] = self.getCheckNum(w_REV[:17])
                                # self.result_signal.emit(f'w_rev:{w_REV}')
                                if not self.is_running:
                                    self.isCancelAllTest = True
                                    return False

                                self.result_signal.emit(f'开始写入选项板REV码：{option_rev}。\n')
                                self.pauseOption()
                                if not self.is_running:
                                    self.isCancelAllTest = True
                                    return False
                                self.CPU_optionPanel(transmitData=w_REV)
                                if not self.isCancelAllTest:
                                    if not self.is_running:
                                        self.isCancelAllTest = True
                                        return False

                                    self.result_signal.emit('写入REV成功！\n')
                                    self.result_signal.emit('开始读取选项板REV码。\n\n')
                                    r_REV = [0xac, 8, 0x00, 0x0f, 0x0e, 0x00, 0x00, 0xf3, 0x3b]
                                    # r_REV[8] = self.getCheckNum(r_REV[:8])
                                    self.pauseOption()
                                    if not self.is_running:
                                        self.isCancelAllTest = True
                                        return False
                                    self.CPU_optionPanel(transmitData=r_REV, param2=option_rev)
                                    if not self.isCancelAllTest:
                                        if not self.is_running:
                                            self.isCancelAllTest = True
                                            return False

                                        self.result_signal.emit('读REV成功！\n\n')
                                        return True
                                    else:
                                        return False
                                else:
                                    return False
                            else:
                                return False
                        else:
                            return False
                    else:
                        return False
                else:
                    return False
            else:
                return False
        except:
            return False

    #选项板测试
    def CPU_optionPanel(self,transmitData:list,baudRate:list=[0x00,0x00,0x00,0x00],
                        param1:list=[0x00,0x00],param2:str=''):
        isThisPass = True
        opType = ''
        try:
            try:
                # 打开串口
                typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
            except serial.SerialException as e:
                if not self.is_running:
                    self.cancelAllTest()
                    self.isPassOp &= False
                    isThisPass = False
                self.messageBox_signal.emit(
                    ['错误警告', f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                                 f'Failed to open serial port: {e}'])
                self.isPassOp = False
                isThisPass = False
                self.cancelAllTest()
            loopStartTime = time.time()
            while True:
                if self.isCancelAllTest:
                    break
                if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                    self.isPassOp &= False
                    isThisPass = False
                    break
                # 发送数据
                typeC_serial.write(bytes(transmitData))
                # 等待0.5s后接收数据
                time.sleep(0.5)
                serial_receiveData = [0 for x in range(40)]
                # 接收数组数据
                data = typeC_serial.read(40)
                serial_receiveData = [hex(i) for i in data]
                for i in range(len(serial_receiveData)):
                    serial_receiveData[i] = int(serial_receiveData[i], 16)
                trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
                if isSendAgain:
                    if not self.is_running:
                        self.cancelAllTest()
                        self.isPassOp &= False
                        isThisPass = False
                        break
                    self.result_signal.emit('未接收到正确数据，再次接收！\n')
                    self.isPassOp &= False
                    isThisPass = False
                    continue
                if dataLen == 3:  # 指令出错
                    self.orderError(trueData[2])
                    self.isPassOp &= False
                    isThisPass = False
                    break
                elif trueData[5] == 0x00:#模拟量
                    if trueData[4] == 0x10:#写模拟量
                        if trueData[7] ==0x00 or trueData[7] ==0x01:#量程/码值
                            if (trueData[8] == 0x01 and trueData[9] == 0x00) or\
                                    (trueData[8] == 0x02 and trueData[9] == 0x00):
                                self.isPassOp &= True
                                isThisPass = True
                            else:
                                self.isPassOp &= False
                                isThisPass = False
                            break
                        elif trueData[7] ==0xF1 or trueData[7] ==0xF2 or trueData[7] ==0xF3:#三码
                            if trueData[8] == 0x00:
                                self.isPassOp &= True
                                isThisPass = True
                            else:
                                self.isPassOp &= False
                                isThisPass = False
                            break
                    elif trueData[4] == 0x0E:#读模拟量
                        if trueData[7] == 0x00 or trueData[7] ==0x01:#量程/码值
                            if (trueData[8] == 0x01 and trueData[9] == 0x00) or \
                                    (trueData[8] == 0x02 and trueData[9] == 0x00):
                                if trueData[7] == 0x00:
                                    self.isPassOp &= True
                                    isThisPass = True
                                elif trueData[7] ==0x01:
                                    r_data = bytes([trueData[10], trueData[11]])
                                    int_r_data = struct.unpack('<h', r_data)[0]
                                    w_data = bytes([param1[0], param1[1]])
                                    int_w_data = struct.unpack('<h', w_data)[0]
                                    if not self.is_running:
                                        self.cancelAllTest()
                                        self.isPassOp &= False
                                        isThisPass = False
                                        break
                                    self.result_signal.emit(f'-----------AQ0004写入值转换比例后的值为{int_w_data}。-----------\n')
                                    self.result_signal.emit(f'-----------模拟量输入通道接收的到的值为{int_r_data}-----------。\n')
                                    # if abs(r_data-w_data)<int(0.001*int_w_data):
                                    if abs(int_r_data - int_w_data) < 40:
                                        self.isPassOp &= True
                                        isThisPass = True
                                    else:
                                        self.isPassOp &= False
                                        isThisPass = False
                            else:
                                self.isPassOp &= False
                                isThisPass = False
                            break
                        elif trueData[7] == 0xF1 or trueData[7] == 0xF2 or trueData[7] == 0xF3:  # 三码
                            code3 = {0xF1:'PN码',0xF2:'SN码',0xF3:'REV码'}
                            if trueData[8] == 0x00:
                                code_PNSNREV = ''
                                codeLen = int(trueData[9])
                                for ll in range(codeLen):
                                    code_PNSNREV += chr(int(trueData[10+ll]))
                                    if int(trueData[10+ll]) != ord(param2[ll]):
                                        self.isPassOp &= False
                                        isThisPass = False
                                        if not self.is_running:
                                            self.cancelAllTest()
                                            self.isPassOp &= False
                                            isThisPass = False
                                            break
                                        self.result_signal.emit(f'{code3[trueData[7]]}写入错误\n\n')
                                        break
                                if not self.is_running:
                                    self.cancelAllTest()
                                    self.isPassOp &= False
                                    isThisPass = False
                                    break
                                self.result_signal.emit(f'读到的{code3[trueData[7]]}为:{code_PNSNREV}\n\n')
                                self.isPassOp &= True
                                isThisPass = True
                            elif trueData[8] == 0x01:
                                self.isPassOp &= False
                                isThisPass = False
                            break
                elif trueData[5] == 0x02:#485
                    if trueData[4] == 0x10:#485写
                        if trueData[7] == 0x00:#485写成功
                            self.isPassOp &= True
                            isThisPass = True
                        elif trueData[7] == 0x01:#485写失败
                            self.isPassOp &= False
                            isThisPass = False
                        break
                    elif trueData[4] == 0x0E:#485读
                        if trueData[7] == 0x00:#485读成功
                            if trueData[8] == baudRate[0] and trueData[9] == baudRate[1]\
                                and trueData[10] == baudRate[2] and trueData[11] == baudRate[3]:
                                self.isPassOp &= True
                                isThisPass = True
                            else:
                                self.isPassOp &= False
                                isThisPass = False
                        elif trueData[7] == 0x00:#485读失败
                            self.isPassOp &= False
                            isThisPass = False
                        break
                elif trueData[5] == 0x03:#232
                    if trueData[4] == 0x10:#232写
                        if trueData[7] == 0x00:#232写成功
                            self.isPassOp &= True
                            isThisPass = True
                        elif trueData[7] == 0x01:#232写失败
                            self.isPassOp &= False
                            isThisPass = False
                        break
                    elif trueData[4] == 0x0E:#232读
                        if trueData[7] == 0x00:#232读成功
                            if trueData[8] == baudRate[0] and trueData[9] == baudRate[1]\
                                and trueData[10] == baudRate[2] and trueData[11] == baudRate[3]:
                                self.isPassOp &= True
                                isThisPass = True
                            else:
                                self.isPassOp &= False
                                isThisPass = False
                        elif trueData[7] == 0x00:#232读失败
                            self.isPassOp &= False
                            isThisPass = False
                        break

                elif trueData[5] == 0xFF:#读选项板型号
                    if trueData[7] == 0x00:
                        if trueData[8] == 0x00:
                            self.messageBox_signal.emit(['操作提示',
                                                         f'选项板型号是否为MA0202？\n'])
                            reply = self.result_queue.get()
                            if reply == QMessageBox.AcceptRole:
                                self.isPassOp &= True
                                isThisPass = True
                                opType = 'MA0202'
                            else:
                                self.isPassOp &= False
                                isThisPass = False
                            break
                        elif trueData[8] == 0x01:
                            self.messageBox_signal.emit(['操作提示',
                                                         f'选项板型号是否为RS-422？\n'])
                            reply = self.result_queue.get()
                            if reply == QMessageBox.AcceptRole:
                                self.isPassOp &= True
                                isThisPass = True
                                opType = '422'
                            else:
                                self.isPassOp &= False
                                isThisPass = False
                            break
                        elif trueData[8] == 0x02:
                            self.messageBox_signal.emit(['操作提示',
                                                         f'选项板型号是否为RS-485？\n'])
                            reply = self.result_queue.get()
                            if reply == QMessageBox.AcceptRole:
                                self.isPassOp &= True
                                isThisPass = True
                                opType = '485'
                            else:
                                self.isPassOp &= False
                                isThisPass = False
                            break
                        elif trueData[8] == 0x03:
                            self.messageBox_signal.emit(['操作提示',
                                                         f'选项板型号是否为RS-232？\n'])
                            reply = self.result_queue.get()
                            if reply == QMessageBox.AcceptRole:
                                self.isPassOp &= True
                                isThisPass = True
                                opType = '232'
                            else:
                                self.isPassOp &= False
                                isThisPass = False
                            break
                        else:
                            if not self.is_running:
                                self.cancelAllTest()
                                self.isPassOp &= False
                                isThisPass = False
                                break
                            self.result_signal.emit('读取选项板类型失败！\n')
                            self.isPassOp &= False
                            isThisPass = False
                        break

                    elif trueData[7] == 0x01:
                        if not self.is_running:
                            self.cancelAllTest()
                            self.isPassOp &= False
                            isThisPass = False
                            break
                        self.result_signal.emit('读取选项板类型失败！\n')
                        self.isPassOp &= False
                        isThisPass = False
                        break
        except:
            if not self.is_running:
                self.cancelAllTest()
                self.isPassOp &= False
                isThisPass = False

            self.result_signal.emit(f"选项板错误信息:\n{traceback.format_exc()}\n")
            self.isPassOp &= False
            isThisPass = False
        finally:
            if not isThisPass:
                self.messageBox_signal.emit(['测试警告', '选项板测试不通过，是否取消后续测试？'])
                reply = self.result_queue.get()
                if reply == QMessageBox.AcceptRole:
                    self.cancelAllTest()
            # 关闭串口
            typeC_serial.close()
            return opType

    # 获取校验位
    def getCheckNum(self, data: list):
        sum = 0
        for d in data:
            sum += d
        checkNum = (~sum) & 0xff
        return checkNum

    def orderError(self, errCode):
        errorCode = {0x80: '消息长度错误', 0x81: '设备错误', 0x82: '测试类型错误',
                     0x83: '操作不支持', 0x84: '参数错误'}
        self.showErrorInf(f'指令发送错误：{hex(errCode)}，{errorCode[errCode]}\n'
                          f'取消剩余测试。\n\n')
        self.isCancelAllTest = True
            
    # 数组元素归零
    def clearList(self, array):
        for i in range(len(array)):
            array[i] = 0x00

    def pause_work(self):
        self.is_pause = True

    def resume_work(self):
        self.is_pause = False

    def pauseOption(self):
        if self.is_pause:
            while True:
                if self.pause_num == 1 and self.is_running:
                    self.pause_num += 1
                    self.result_signal.emit(self.HORIZONTAL_LINE + '暂停中…………' + self.HORIZONTAL_LINE)
                if not self.is_pause:
                    self.pause_num = 1
                    return False

    def stop_work(self):
        self.resume_work()
        self.is_running = False
        # ·
        self.label_signal.emit(['fail', '测试停止'])

    def showErrorInf(self, Inf):
        self.pauseOption()
        if not self.is_running:
            self.isCancelAllTest = True
            return
        self.result_signal.emit(f"{Inf}程序出现问题\n")
        # 捕获异常并输出详细的错误信息
        self.result_signal.emit(f"ErrorInf:\n{traceback.format_exc()}\n")
        self.messageBox_signal.emit(['错误警告', (f"ErrorInf:\n{traceback.format_exc()}\n")])
        
    def clearList(self, array):
        for i in range(len(array)):
            array[i] = 0x00