# -*- coding: utf-8 -*-

import time

from PyQt5.QtCore import QThread, pyqtSignal
from main_logic import *
import threading
from CAN_option import *
import otherOption


class DIDOThread(QObject):
    result_signal = pyqtSignal(str)
    item_signal = pyqtSignal(list)
    pass_signal = pyqtSignal(bool)
    # RunErr_signal = pyqtSignal(int)
    # CANRunErr_signal = pyqtSignal(int)
    messageBox_signal = pyqtSignal(list)
    # excel_signal = pyqtSignal(list)
    allFinished_signal = pyqtSignal()
    label_signal = pyqtSignal(list)
    saveExcel_signal = pyqtSignal(list)
    print_signal = pyqtSignal(list)

    HORIZONTAL_LINE = "\n------------------------------------------------------------------------------------------------------------\n\n"
    m_arrayTestData = [[0x01, 0x00, 0x00, 0x00], [0x02, 0x00, 0x00, 0x00], [0x04, 0x00, 0x00, 0x00],
                       [0x08, 0x00, 0x00, 0x00], [0x10, 0x00, 0x00, 0x00], [0x20, 0x00, 0x00, 0x00],
                       [0x40, 0x00, 0x00, 0x00],
                       [0x80, 0x00, 0x00, 0x00], [0x00, 0x01, 0x00, 0x00], [0x00, 0x02, 0x00, 0x00],
                       [0x00, 0x04, 0x00, 0x00], [0x00, 0x08, 0x00, 0x00], [0x00, 0x10, 0x00, 0x00],
                       [0x00, 0x20, 0x00, 0x00], [0x00, 0x40, 0x00, 0x00], [0x00, 0x80, 0x00, 0x00],
                       [0x00, 0x00, 0x01, 0x00], [0x00, 0x00, 0x02, 0x00], [0x00, 0x00, 0x04, 0x00],
                       [0x00, 0x00, 0x08, 0x00], [0x00, 0x00, 0x10, 0x00], [0x00, 0x00, 0x20, 0x00],
                       [0x00, 0x00, 0x40, 0x00], [0x00, 0x00, 0x80, 0x00], [0x00, 0x00, 0x00, 0x01],
                       [0x00, 0x00, 0x00, 0x02], [0x00, 0x00, 0x00, 0x04], [0x00, 0x00, 0x00, 0x08],
                       [0x00, 0x00, 0x00, 0x10], [0x00, 0x00, 0x00, 0x20], [0x00, 0x00, 0x00, 0x40],
                       [0x00, 0x00, 0x00, 0x80], [0xff, 0xff, 0xff, 0xff], [0x00, 0x00, 0x00, 0x00],
                       [0xaa, 0xaa, 0xaa, 0xaa],[0x55, 0x55, 0x55, 0x55]
                       ]

    # 接收的数据
    ubyte_array_receive = c_ubyte * 8
    m_receiveData = ubyte_array_receive(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
    TIME_OUT = 1000  # ms
    interval = 500  # ms

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

    volReceValue = [0, 0, 0, 0, 0]
    curReceValue = [0, 0, 0, 0, 0]
    volPrecision = [0, 0, 0, 0, 0]
    curPrecision = [0, 0, 0, 0, 0]

    # AO测试是否通过
    isAOPassTest = True
    isAOVolPass = True
    isAOCurPass = True

    isLEDRunOK = True
    isLEDErrOK = True
    isLEDPass = True

    isLEDCANRunOK = True
    isLEDCANErrOK = True
    isLEDCANPass = True

    # 发送的数据
    ubyte_array_transmit = c_ubyte * 8
    m_transmitData = ubyte_array_transmit(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
    # 主副线程弹窗结果
    result_queue = 0

    CAN_errorLED = True
    CAN_runLED = True
    errorLED = True
    runLED = True

    def __init__(self, inf_DIDOlist: list, result_queue, appearance,testFlage):
        super().__init__()
        self.testFlage = testFlage
        self.result_queue = result_queue
        self.is_running = True
        self.is_pause = False
        self.appearance = appearance
        self.mTable = inf_DIDOlist[0][0]
        self.module_1 = inf_DIDOlist[0][1]
        self.module_2 = inf_DIDOlist[0][2]
        self.testNum = inf_DIDOlist[0][3]

        # 获取产品信息
        self.module_type = inf_DIDOlist[1][0]
        self.module_pn = inf_DIDOlist[1][1]
        self.module_sn = inf_DIDOlist[1][2]
        self.module_rev = inf_DIDOlist[1][3]
        self.m_Channels = int(inf_DIDOlist[1][4])
        self.DIDO_Channels = self.m_Channels

        # 获取CAN地址
        self.CANAddr_DI = inf_DIDOlist[2][0]
        self.CANAddr_DO = inf_DIDOlist[2][1]
        # self.CANAddr_relay = inf_DIDOlist[2][2]

        # 获取附加参数
        self.baud_rate = inf_DIDOlist[3][0]
        self.waiting_time = inf_DIDOlist[3][1]
        self.loop_num = inf_DIDOlist[3][2]
        self.saveDir = inf_DIDOlist[3][3]

        # # 获取标定信息
        # self.isCalibrate = inf_AOlist[4][0]
        # self.isCalibrateVol = inf_AOlist[4][1]
        # self.isCalibrateCur = inf_AOlist[4][2]
        # self.oneOrAll = inf_AOlist[4][3]
        #
        # # 获取检测信息
        # self.isTest = inf_AOlist[5][0]
        # self.isAOTestVol = inf_AOlist[5][1]
        # self.isAOTestCur = inf_AOlist[5][2]
        self.isTestCANRunErr = inf_DIDOlist[4][0]
        self.isTestRunErr = inf_DIDOlist[4][1]
        # print(f'inf_AOlist[5] = {inf_AOlist[5]}')
        self.errorNum = 0
        self.errorInf = ''
        self.pause_num = 1
        errorNum = 0
        # DO通道数据初始化
        self.DO_channelData = 0
        # DI通道数据初始化
        self.DI_channelData = 0

        self.isDIPassTest = True
        self.isDOPassTest = True
        self.isPassOdd = True
        self.isPassEven = True
    def DIDOOption(self):
        isExcel = True
        if self.testFlage == 'DI':
            self.testDI()
        elif self.testFlage == 'DO':
            self.testDO()

        if isExcel:
            self.result_signal.emit('开始生成校准校验表…………' + self.HORIZONTAL_LINE)
            code_array = [self.module_pn, self.module_sn, self.module_rev]

            if self.testFlage == 'DI':
                station_array = [self.isDIPassTest, '', '']
                excel_bool, book, sheet, self.DI_row = otherOption.generateExcel(code_array,
                                                                                 station_array,
                                                                                 self.DIDO_Channels, 'DI')
                if not excel_bool:
                    self.result_signal.emit('DI校准校验表生成出错！请检查代码！' + self.HORIZONTAL_LINE)
                else:
                    self.fillInDIData(self.isDIPassTest, book, sheet)
                    self.result_signal.emit('生成DI校准校验表成功' + self.HORIZONTAL_LINE)
            elif self.testFlage == 'DO':
                station_array = [self.isDOPassTest, '', '']
                excel_bool, book, sheet, self.DO_row = otherOption.generateExcel(code_array,
                                                                                 station_array,
                                                                                 self.DIDO_Channels, 'DO')
                if not excel_bool:
                    self.result_signal.emit('DO校准校验表生成出错！请检查代码！' + self.HORIZONTAL_LINE)
                else:
                    self.fillInDOData(self.isDOPassTest, book, sheet)
                    self.result_signal.emit('生成DO校准校验表成功' + self.HORIZONTAL_LINE)
        elif not isExcel:
            self.result_signal.emit('测试停止，校准校验表生成失败…………' + self.HORIZONTAL_LINE)

        self.allFinished_signal.emit()
        self.pass_signal.emit(True)

    def CAN_init(self):
        CAN_option.connect(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX, CAN_option.CAN_INDEX)

        CAN_option.init(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX, CAN_option.CAN_INDEX, init_config)

        CAN_option.start(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX, CAN_option.CAN_INDEX)

    def testDI(self):
        DI_startTime = time.time()
        if self.isTestRunErr or self.isTestCANRunErr:
            # 进入指示灯测试模式
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("--------------进入 LED TEST模式-------------\n\n")
            print("--------------进入 LED TEST模式-------------")
            # self.channelZero()
            self.m_transmitData[0] = 0x23
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x00
            self.m_transmitData[4] = 0x53
            self.m_transmitData[5] = 0x54
            self.m_transmitData[6] = 0x41
            self.m_transmitData[7] = 0x52
            isLEDTest, whatEver = CAN_option.transmitCAN((0x600 + int(self.CANAddr_DI)), self.m_transmitData)
            if isLEDTest:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit("-----------进入 LED TEST 模式成功-----------\n")
                print("-----------进入 LED TEST 模式成功-----------" + self.HORIZONTAL_LINE)

                if self.isTestRunErr:
                    if not self.testRunErr(int(self.CANAddr_DI)):
                        self.pass_signal.emit(False)
                if self.isTestCANRunErr:
                    if not self.testCANRunErr(int(self.CANAddr_DI)):
                        self.pass_signal.emit(False)

                # 退出 LED TEST 模式
                self.clearList(self.m_transmitData)
                self.m_transmitData[0] = 0x23
                self.m_transmitData[1] = 0xf6
                self.m_transmitData[2] = 0x5f
                self.m_transmitData[3] = 0x00
                self.m_transmitData[4] = 0x45
                self.m_transmitData[5] = 0x58
                self.m_transmitData[6] = 0x49
                self.m_transmitData[7] = 0x54
                bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + self.CANAddr_DI),
                                                                       self.m_transmitData)
                if bool_transmit:
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    self.result_signal.emit("成功退出 LED TEST 模式。\n" + self.HORIZONTAL_LINE)
                    print("成功退出 LED TEST 模式\n" + self.HORIZONTAL_LINE)
                else:
                    self.result_signal.emit("退出 LED TEST 模式失败！\n" + self.HORIZONTAL_LINE)
            else:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit("-----------进入 LED TEST 模式失败-----------\n\n")
                print("-----------进入 LED TEST 模式失败-----------\n\n")
                self.item_signal.emit([1, 2, 2, '进入模式失败'])
                self.item_signal.emit([2, 2, 2, '进入模式失败'])

        self.item_signal.emit([5, 1, 0, ''])
        self.DIDataCheck = [True for i in range(32)]
        for i in range(self.loop_num):
            # if not self.isStop():
            #     return
            self.result_signal.emit(f"第{i+1}次循环"+self.HORIZONTAL_LINE)
            CAN_option.Can_DLL.VCI_ClearBuffer(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX, CAN_option.CAN_INDEX)
            time.sleep(self.interval / 1000)  # s
            # 通道全亮
            self.testByIndex(32,'DI')
            time.sleep(self.interval / 1000)
            # 偶数通道全亮
            self.testByIndex(34,'DI')
            time.sleep(self.interval / 1000)
            # 奇数通道全亮
            self.testByIndex(35,'DI')
            time.sleep(self.interval / 1000)
            for j in range(self.m_Channels):
                # if not self.isStop():
                #     return
                self.testByIndex(j,'DI')
                self.DI_channelData |= self.m_receiveData[0]
                self.DI_channelData |= self.m_receiveData[1] << 8
                self.DI_channelData |= self.m_receiveData[2] << 16
                self.DI_channelData |= self.m_receiveData[3] << 24
                time.sleep(self.interval / 1000)
            self.testByIndex(33,'DI')
        DI_endTime = time.time()
        DI_testTime = round((DI_endTime - DI_startTime),3)
        if self.isDIPassTest == False:
            self.messageBox_signal.emit(['DI通道指示灯结果', f'DI通道指示灯未全部正常显示！'])

            self.result_signal.emit(self.HORIZONTAL_LINE + 'DI通道指示灯测试未通过\n' + self.HORIZONTAL_LINE)
            self.item_signal.emit([5,2,2,f'{DI_testTime}'])
        elif self.isDIPassTest == True:
            # if not self.isStop():
            #
            #     return
            # reply = QMessageBox.about(None, 'DI通道指示灯结果', f'DI通道指示灯全部正常显示！')
            self.messageBox_signal.emit(['DI通道指示灯结果', f'DI通道指示灯全部正常显示！'])
            self.result_signal.emit(self.HORIZONTAL_LINE + 'DI通道指示灯测试全通过\n' + self.HORIZONTAL_LINE)
            self.item_signal.emit([5, 2, 1, f'{DI_testTime}'])
        # self.endOfTest()

    def testDO(self):
        DO_startTime = time.time()
        if self.isTestRunErr or self.isTestCANRunErr:
            # 进入指示灯测试模式
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("--------------进入 LED TEST模式-------------\n\n")
            print("--------------进入 LED TEST模式-------------")
            # self.channelZero()
            self.m_transmitData[0] = 0x23
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x00
            self.m_transmitData[4] = 0x53
            self.m_transmitData[5] = 0x54
            self.m_transmitData[6] = 0x41
            self.m_transmitData[7] = 0x52
            isLEDTest, whatEver = CAN_option.transmitCAN((0x600 + int(self.CANAddr_DO)), self.m_transmitData)
            if isLEDTest:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit("-----------进入 LED TEST 模式成功-----------\n")
                print("-----------进入 LED TEST 模式成功-----------" + self.HORIZONTAL_LINE)

                if self.isTestRunErr:
                    if not self.testRunErr(int(self.CANAddr_DO)):
                        self.pass_signal.emit(False)
                if self.isTestCANRunErr:
                    if not self.testCANRunErr(int(self.CANAddr_DO)):
                        self.pass_signal.emit(False)

                #退出 LED TEST 模式
                self.clearList(self.m_transmitData)
                self.m_transmitData[0] = 0x23
                self.m_transmitData[1] = 0xf6
                self.m_transmitData[2] = 0x5f
                self.m_transmitData[3] = 0x00
                self.m_transmitData[4] = 0x45
                self.m_transmitData[5] = 0x58
                self.m_transmitData[6] = 0x49
                self.m_transmitData[7] = 0x54
                bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + self.CANAddr_DO), self.m_transmitData)
                if bool_transmit:
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    self.result_signal.emit("成功退出 LED TEST 模式。\n" + self.HORIZONTAL_LINE)
                    print("成功退出 LED TEST 模式\n" + self.HORIZONTAL_LINE)
                else:
                    self.result_signal.emit("退出 LED TEST 模式失败！\n" + self.HORIZONTAL_LINE)
            else:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit("-----------进入 LED TEST 模式失败-----------\n\n")
                print("-----------进入 LED TEST 模式失败-----------\n\n")
                self.item_signal.emit([1, 2, 2, '进入模式失败'])
                self.item_signal.emit([2, 2, 2, '进入模式失败'])

        self.item_signal.emit([5, 1, 0, ''])
        self.DODataCheck = [True for i in range(32)]#预留了最多32个通道的测试结果
        for i in range(self.loop_num):
            # if not self.isStop():
            #     return
            CAN_option.Can_DLL.VCI_ClearBuffer(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX, CAN_option.CAN_INDEX)
            time.sleep(self.interval / 1000)
            # 通道全亮
            self.testByIndex(32,'DO')
            time.sleep(self.interval / 1000)
            # 偶数通道全亮
            self.testByIndex(34,'DO')
            time.sleep(self.interval / 1000)
            # 奇数通道全亮
            self.testByIndex(35,'DO')
            time.sleep(self.interval / 1000)
            for j in range(self.m_Channels):
                # if not self.isStop():
                #     return
                self.testByIndex(j,'DO')
                self.DO_channelData |= self.m_receiveData[0]
                self.DO_channelData |= self.m_receiveData[1] << 8
                self.DO_channelData |= self.m_receiveData[2] << 16
                self.DO_channelData |= self.m_receiveData[3] << 24
                time.sleep(self.interval / 1000)
            self.testByIndex(33,'DO')
        DO_endTime = time.time()
        DO_testTime = round((DO_endTime - DO_startTime), 3)
        if self.isDOPassTest == False:
            # reply = QMessageBox.warning(None, 'DO通道指示灯结果', 'DO通道指示灯未全部正常显示！',
            #                             QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            self.messageBox_signal.emit(['DO通道指示灯结果', f'DO通道指示灯未全部正常显示！'])
            self.result_signal.emit(self.HORIZONTAL_LINE + 'DO通道指示灯测试未通过\n' + self.HORIZONTAL_LINE)
            self.item_signal.emit([5, 2, 2, f'{DO_testTime}'])
        elif self.isDOPassTest == True:
            # reply = QMessageBox.about(None, 'DO通道指示灯结果', 'DO通道指示灯全部正常显示！')
            self.messageBox_signal.emit(['DO通道指示灯结果', f'DO通道指示灯全部正常显示！'])
            self.result_signal.emit(self.HORIZONTAL_LINE + 'DO通道指示灯测试全通过\n' + self.HORIZONTAL_LINE)
            self.item_signal.emit([5, 2, 1, f'{DO_testTime}'])

        # self.endOfTest()


    # def calibrateByIndex(self, index):
    #     # if not self.isStop():
    #     #     return
    #     self.sendTestDataToDO(index)
    #     print(f'index={index}')
    #     if index == 32:
    #         self.result_signal.emit(
    #             f'1.发送的数据：{hex(self.m_transmitData[0])}  {hex(self.m_transmitData[1])}  '
    #             f'{hex(self.m_transmitData[2])}  '
    #             f'{hex(self.m_transmitData[3])}\n\n')
    #     elif index == 33:
    #         self.result_signal.emit(
    #             f'{self.m_Channels + 4}.发送的数据：{hex(self.m_transmitData[0])}  {hex(self.m_transmitData[1])}  '
    #             f'{hex(self.m_transmitData[2])}  '
    #             f'{hex(self.m_transmitData[3])}\n\n')
    #     elif index == 34:
    #         self.result_signal.emit(
    #             f'2.发送的数据：{hex(self.m_transmitData[0])}  {hex(self.m_transmitData[1])}  '
    #             f'{hex(self.m_transmitData[2])}  '
    #             f'{hex(self.m_transmitData[3])}\n\n')
    #     elif index == 35:
    #         self.result_signal.emit(
    #             f'3.发送的数据：{hex(self.m_transmitData[0])}  {hex(self.m_transmitData[1])}  '
    #             f'{hex(self.m_transmitData[2])}  '
    #             f'{hex(self.m_transmitData[3])}\n\n')
    #     else:
    #         self.result_signal.emit(
    #             f'{index + 4}.发送的数据：{hex(self.m_transmitData[0])}  {hex(self.m_transmitData[1])}  '
    #             f'{hex(self.m_transmitData[2])}  '
    #             f'{hex(self.m_transmitData[3])}\n\n')
    #
    #     now_time = time.time()
    #     while True:
    #         # if not self.isStop():
    #         #     return
    #         if (time.time() - now_time) * 1000 > self.TIME_OUT:
    #             self.result_signal.emit(
    #                 f'  接收的数据：{hex(self.m_receiveData[0])}  {hex(self.m_receiveData[1])}  '
    #                 f'{hex(self.m_receiveData[2])}  '
    #                 f'{hex(self.m_receiveData[3])}  收发不一致！\n\n')
    #             print(
    #                 f'接收的数据：{hex(self.m_receiveData[0])}  {hex(self.m_receiveData[1])}  '
    #                 f'{hex(self.m_receiveData[2])}  '
    #                 f'{hex(self.m_receiveData[3])}  收发不一致！\n\n')
    #             if index != 32 and index != 33:

    #             print('收发不一致！\n\n')
    #             self.isDIPassTest = False
    #             # reply = QMessageBox.about(None, '警告', '检测到收发不一致！')
    #             self.messageBox_signal.emit(['警告', '检测到收发不一致！'])
    #             reply = self.result_queue.get()
    #             if reply == QMessageBox.Yes:
    #                 break
    #             else:
    #                  break
    #         # 清理接收缓存区
    #         self.clearList(self.m_receiveData)
    #         bool_receive, self.m_can_obj = CAN_option.receiveCANbyID((0x180 + self.CANAddr_DI), self.TIME_OUT)
    #         self.m_receiveData = self.m_can_obj.Data
    #         if bool_receive == False:
    #             self.result_signal.emit('\n  接收数据超时！\n\n')
    #             print('接收数据超时！')
    #             self.isDIPassTest = False
    #             return
    #         elif self.m_transmitData[0] == self.m_receiveData[0] and \
    #                 self.m_transmitData[1] == self.m_receiveData[1] and \
    #                 self.m_transmitData[2] == self.m_receiveData[2] and \
    #                 self.m_transmitData[3] == self.m_receiveData[3]:
    #             self.result_signal.emit(f'  接收的数据：{hex(self.m_receiveData[0])}  {hex(self.m_receiveData[1])}  '
    #                          f'{hex(self.m_receiveData[2])}  {hex(self.m_receiveData[3])}  收发一致\n\n')
    #             print(f'接收的数据：{hex(self.m_receiveData[0])}  {hex(self.m_receiveData[1])}  '
    #                   f'{hex(self.m_receiveData[2])}  {hex(self.m_receiveData[3])}  收发一致\n\n')
    #             return

    def testByIndex(self, index,type:str):
        self.sendTestDataToDO(index)
        print(f'index={index}')
        if index == 32:
            self.result_signal.emit(
                f'{1}.发送的数据：{hex(self.m_transmitData[0])}  {hex(self.m_transmitData[1])}  {hex(self.m_transmitData[2]) }'
                f'{hex(self.m_transmitData[3])}\n\n')
        elif index == 33:
            self.result_signal.emit(
                f'{self.m_Channels + 2}.发送的数据：{hex(self.m_transmitData[0])}  '
                f'{hex(self.m_transmitData[1])}  {hex(self.m_transmitData[2])}  '
                f'{hex(self.m_transmitData[3])}\n\n')
        else:
            self.result_signal.emit(f'{index + 2}.发送的数据：{hex(self.m_transmitData[0])}  '
                         f'{hex(self.m_transmitData[1])}  {hex(self.m_transmitData[2])}  '
                         f'{hex(self.m_transmitData[3])}\n\n')

        now_time = time.time()
        while True:
            # if not self.isStop():
            #     return
            if (time.time() - now_time) * 1000 > self.TIME_OUT:

                self.result_signal.emit(f'  接收的数据：{hex(self.m_receiveData[0])}  '
                             f'{hex(self.m_receiveData[1])}  {hex(self.m_receiveData[2])}  '
                             f'{hex(self.m_receiveData[3])}  收发不一致！\n\n')

                print(f'接收的数据：{hex(self.m_receiveData[0])}  {hex(self.m_receiveData[1])}  '
                      f'{hex(self.m_receiveData[2])}  '
                      f'{hex(self.m_receiveData[3])}  收发不一致！\n\n')

                if type == 'DO':
                    if index != 32 and index != 33 and index != 34 and index != 35:
                        self.DODataCheck[index] = False
                    if index == 34:
                        self.isPassEven = False
                    if index == 35:
                        self.isPassOdd = False
                    self.isDOPassTest = False
                elif type == 'DI':
                    if index != 32 and index != 33 and index != 34 and index != 35:
                        self.DIDataCheck[index] = False
                    if index == 34:
                        self.isPassEven = False
                    if index == 35:
                        self.isPassOdd = False
                    self.isDIPassTest = False

                # reply = QMessageBox.about(None, '警告', '检测到收发不一致！')
                self.messageBox_signal.emit(['警告', '检测到收发不一致！'])
                reply = self.result_queue.get()
                if reply == QMessageBox.Yes:
                    break
                else:
                     break
            # 清理接收缓存区
            self.clearList(self.m_receiveData)
            bool_receive, self.m_can_obj = CAN_option.receiveCANbyID((0x180 + self.CANAddr_DI), self.TIME_OUT)
            self.m_receiveData = self.m_can_obj.Data
            if bool_receive == False:
                self.result_signal.emit('\n  接收数据超时！\n\n')
                print('接收数据超时！')
                if type == 'DO':
                    self.isDOPassTest = False
                    self.DODataCheck[index] = False
                if type == 'DI':
                    self.isDIPassTest = False
                    self.DIDataCheck[index] = False
                return
            elif self.m_transmitData[0] == self.m_receiveData[0] and \
                    self.m_transmitData[1] == self.m_receiveData[1] and \
                    self.m_transmitData[2] == self.m_receiveData[2] and \
                    self.m_transmitData[3] == self.m_receiveData[3]:
                self.result_signal.emit(f'  接收的数据：{hex(self.m_receiveData[0])}  {hex(self.m_receiveData[1])}  '
                             f'{hex(self.m_receiveData[2])}  {hex(self.m_receiveData[3])}  收发一致\n\n')
                print(f'接收的数据：{hex(self.m_receiveData[0])}  {hex(self.m_receiveData[1])}  '
                      f'{hex(self.m_receiveData[2])}  {hex(self.m_receiveData[3])}  收发一致\n\n')
                return

    def clearList(self, array):
        for i in range(len(array)):
            array[i] = 0x00

    def sendTestDataToDO(self, index):
        if index == 32 and self.m_Channels <= 16:
            if self.m_Channels == 4:
                self.m_transmitData[0] = 0x0f
                self.m_transmitData[1] = 0x00
                self.m_transmitData[2] = 0x00
                self.m_transmitData[3] = 0x00
            elif self.m_Channels == 8:
                self.m_transmitData[0] = 0xff
                self.m_transmitData[1] = 0x00
                self.m_transmitData[2] = 0x00
                self.m_transmitData[3] = 0x00
            elif self.m_Channels == 16:
                self.m_transmitData[0] = 0xff
                self.m_transmitData[1] = 0xff
                self.m_transmitData[2] = 0x00
                self.m_transmitData[3] = 0x00
        elif index == 34 and self.m_Channels <= 16:
            if self.m_Channels == 4:
                self.m_transmitData[0] = 0x0a
                self.m_transmitData[1] = 0x00
                self.m_transmitData[2] = 0x00
                self.m_transmitData[3] = 0x00
            elif self.m_Channels == 8:
                self.m_transmitData[0] = 0xaa
                self.m_transmitData[1] = 0x00
                self.m_transmitData[2] = 0x00
                self.m_transmitData[3] = 0x00
            elif self.m_Channels == 16:
                self.m_transmitData[0] = 0xaa
                self.m_transmitData[1] = 0xaa
                self.m_transmitData[2] = 0x00
                self.m_transmitData[3] = 0x00
        elif index == 35 and self.m_Channels <= 16:
            if self.m_Channels == 4:
                self.m_transmitData[0] = 0x05
                self.m_transmitData[1] = 0x00
                self.m_transmitData[2] = 0x00
                self.m_transmitData[3] = 0x00
            elif self.m_Channels == 8:
                self.m_transmitData[0] = 0x55
                self.m_transmitData[1] = 0x00
                self.m_transmitData[2] = 0x00
                self.m_transmitData[3] = 0x00
            elif self.m_Channels == 16:
                self.m_transmitData[0] = 0x55
                self.m_transmitData[1] = 0x55
                self.m_transmitData[2] = 0x00
                self.m_transmitData[3] = 0x00
        else:
            self.m_transmitData[0] = self.m_arrayTestData[index][0]
            self.m_transmitData[1] = self.m_arrayTestData[index][1]
            self.m_transmitData[2] = self.m_arrayTestData[index][2]
            self.m_transmitData[3] = self.m_arrayTestData[index][3]
        CAN_option.transmitCAN(0x200 + self.CANAddr_DO, self.m_transmitData)

    def testRunErr(self, addr):
        self.testNum -= 1

        self.isLEDRunOK = True

        self.isLEDErrOK = True
        self.isLEDPass = True
        self.pauseOption()
        if not self.is_running:
            return False
        self.item_signal.emit([1, 1, 0, ''])
        # mTable.item(1, i).setBackground(QtGui.QColor(255, 255, 0))
        # if i == 1:
        #     mTable.item(1, i).setText('正在检测')
        runStart_time = time.time()

        self.pauseOption()
        if not self.is_running:
            return False
        self.result_signal.emit("1.进行 LED RUN 测试\n\n")
        print("1.进行 LED RUN 测试\n\n")
        self.clearList(self.m_transmitData)
        self.m_transmitData[0] = 0x2f
        self.m_transmitData[1] = 0xf6
        self.m_transmitData[2] = 0x5f
        self.m_transmitData[3] = 0x01
        self.m_transmitData[4] = 0x01
        bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData)
        runEnd_time = time.time()
        runTest_time = round(runEnd_time - runStart_time, 2)
        time.sleep(0.5)
        # reply = QMessageBox.question(None, '检测RUN &ERROR', 'RUN指示灯是否点亮？',
        #                              QMessageBox.Yes | QMessageBox.No,
        #                              QMessageBox.Yes)
        self.messageBox_signal.emit(['检测RUN &ERROR', 'RUN指示灯是否点亮？'])
        reply = self.result_queue.get()
        if reply == QMessageBox.Yes:
            self.runLED = True
            # for i in range(4):
            #     item = mTable.item(1, i)
            #     item.setBackground(QtGui.QColor(0, 255, 0))
            #     item.setTextAlignment(QtCore.Qt.AlignCenter)
            #     if i == 1:
            #         item.setText('检测完成')
            #     if i == 2:
            #         item.setText('通过')
            #     if i == 3:
            #         item.setText(f'{runTest_time}')
            self.pauseOption()
            if not self.is_running:
                return False
            self.item_signal.emit([1, 2, 1, f'{runTest_time}'])
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("LED RUN 测试通过\n关闭LED RUN\n" + self.HORIZONTAL_LINE)
            print("LED RUN 测试通过\n关闭LED RUN\n" + self.HORIZONTAL_LINE)
            self.clearList(self.m_transmitData)
            self.m_transmitData[0] = 0x2f
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x01
            self.m_transmitData[4] = 0x00
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData)
            time.sleep(1)
            self.isLEDRunOK = True
            # self.result_signal.emit(f'self.isLEDRunOK:{self.isLEDRunOK}')
        elif reply == QMessageBox.No:
            self.runLED = False
            self.pauseOption()
            if not self.is_running:
                return False
            self.item_signal.emit([1, 2, 2, f'{runTest_time}'])
            # for i in range(4):
            #     item = mTable.item(1, i)
            #     item.setBackground(QtGui.QColor(255, 0, 0))
            #     item.setTextAlignment(QtCore.Qt.AlignCenter)
            #     item.setForeground(QtGui.QColor(255, 255, 255))
            #     if i == 1:
            #         item.setText('检测完成')
            #     if i == 2:
            #         item.setText('未通过')
            #     if i == 3:
            #         item.setText(f'{runTest_time}')
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("LED RUN 测试不通过\n" + self.HORIZONTAL_LINE)
            print("LED RUN 测试不通过\n" + self.HORIZONTAL_LINE)
            self.isLEDRunOK = False
            # self.result_signal.emit(f'self.isLEDRunOK:{self.isLEDRunOK}')
        self.isLEDPass = self.isLEDPass & self.isLEDRunOK

        errorStart_time = time.time()
        self.pauseOption()
        if not self.is_running:
            return False
        self.result_signal.emit("2.进行 LED ERROR 测试\n\n")
        print("2.进行 LED ERROR 测试\n\n")
        self.pauseOption()
        if not self.is_running:
            return False
        self.item_signal.emit([2, 1, 0, ''])
        # for i in range(4):
        #     mTable.item(2, i).setBackground(QtGui.QColor(255, 255, 0))
        #     if i == 1:
        #         mTable.item(2, i).setText('正在检测')
        #         mTable.item(2, i).setTextAlignment(QtCore.Qt.AlignCenter)

        self.clearList(self.m_transmitData)
        self.m_transmitData[0] = 0x2f
        self.m_transmitData[1] = 0xf6
        self.m_transmitData[2] = 0x5f
        self.m_transmitData[3] = 0x02
        self.m_transmitData[4] = 0x01

        bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData)
        errorEnd_time = time.time()
        errorTest_time = round(errorEnd_time - errorStart_time, 2)
        time.sleep(0.5)
        self.messageBox_signal.emit(['检测RUN &ERROR', 'ERROR指示灯是否点亮？'])
        reply = self.result_queue.get()
        # reply = QMessageBox.question(None, '检测RUN &ERROR', 'ERROR指示灯是否点亮？',
        #                              QMessageBox.Yes | QMessageBox.No,
        #                              QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            self.errorLED = True
            # for i in range(4):
            #     item = mTable.item(2, i)
            #     item.setBackground(QtGui.QColor(0, 255, 0))
            #     item.setTextAlignment(QtCore.Qt.AlignCenter)
            #     if i == 1:
            #         item.setText('检测完成')
            #     if i == 2:
            #         item.setText('通过')
            #     if i == 3:
            #         item.setText(f'{errorTest_time}')
            # self.itemOperation(mTable,2,2,1,errorTest_time)
            self.pauseOption()
            if not self.is_running:
                return False
            self.item_signal.emit([2, 2, 1, f'{errorTest_time}'])
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("LED ERROR 测试通过\n关闭LED ERROR\n" + self.HORIZONTAL_LINE)
            print("LED ERROR 测试通过\n关闭LED ERROR\n" + self.HORIZONTAL_LINE)
            self.clearList(self.m_transmitData)
            self.m_transmitData[0] = 0x2f
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x02
            self.m_transmitData[4] = 0x00
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData)
            time.sleep(1)
            self.isLEDErrOK = True
        elif reply == QMessageBox.No:
            self.errorLED = False
            self.pauseOption()
            if not self.is_running:
                return False
            self.item_signal.emit([2, 2, 2, f'{errorTest_time}'])
            # self.itemOself.pauseOption()peration(mTable, 2, 2, 2, errorTest_time)
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("LED ERROR 测试不通过\n" + self.HORIZONTAL_LINE)
            print("LED ERROR 测试不通过\n" + self.HORIZONTAL_LINE)
            self.isLEDErrOK = False
        self.isLEDPass = self.isLEDPass & self.isLEDErrOK

        return True

    def testCANRunErr(self, addr):
        self.testNum -= 1
        # 关闭CAN设备
        # CAN_option.close(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX)
        # 启动CAN设备并打开CAN通道
        # self.can_start()
        isLEDCANRunOK = True
        isLEDCANErrOK = True
        # reply = QMessageBox.question(None, '检测CAN_RUN &CAN_ERROR', '是否开始进行CAN_RUN 和CAN_ERROR 检测？',
        #                              QMessageBox.Yes | QMessageBox.No,
        #                              QMessageBox.Yes)
        # if self.tabIndex == 0:
        #     mTable = self.tableWidget_DIDO
        # elif self.tabIndex == 1:
        #     mTable = self.tableWidget_AI
        # elif self.tabIndex == 2:
        #     mTable = self.tableWidget_AO
        # if reply == QMessageBox.Yes:
        CANRunStart_time = time.time()
        self.pauseOption()
        if not self.is_running:
            return False
        self.item_signal.emit([3, 1, 0, ''])
        # 进入指示灯测试模式
        self.pauseOption()
        if not self.is_running:
            return False
        # self.result_signal.emit("--------------进入 LED TEST模式-------------\n\n")
        # print("--------------进入 LED TEST模式-------------")
        # if not self.channelZero():
        #     return False
        self.m_transmitData[0] = 0x23
        self.m_transmitData[1] = 0xf6
        self.m_transmitData[2] = 0x5f
        self.m_transmitData[3] = 0x00
        self.m_transmitData[4] = 0x53
        self.m_transmitData[5] = 0x54
        self.m_transmitData[6] = 0x41
        self.m_transmitData[7] = 0x52
        print(f'{self.module_2}地址:{0x600 + addr}')
        isLEDTest, whatEver = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData)
        if isLEDTest:
            self.pauseOption()
            if not self.is_running:
                return False
            # self.result_signal.emit("-----------进入 LED TEST 模式成功-----------\n")
            # print("-----------进入 LED TEST 模式成功-----------" + self.HORIZONTAL_LINE)
            # self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("1.进行 LED CAN_RUN 测试\n\n")
            print("1.进行 LED CAN_RUN 测试\n\n")
            self.clearList(self.m_transmitData)
            self.m_transmitData[0] = 0x2f
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x03
            self.m_transmitData[4] = 0x01
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData)
            CANRunEnd_time = time.time()
            CANRunTest_time = round(CANRunEnd_time - CANRunStart_time, 2)
            time.sleep(0.5)
            self.messageBox_signal.emit(['检测RUN &ERROR', 'CAN_RUN指示灯是否点亮？'])
            reply = self.result_queue.get()
            # reply = QMessageBox.question(None, '检测CAN_RUN &CAN_ERROR', '指示灯是否点亮？',
            #                              QMessageBox.Yes | QMessageBox.No,
            #                              QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                self.CAN_runLED = True
                self.pauseOption()
                if not self.is_running:
                    return False
                self.item_signal.emit([3, 2, 1, CANRunTest_time])
                # self.itemOperation(mTable,3,2,1,CANRunTest_time)
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit("LED CAN_RUN 测试通过\n关闭LED CAN_RUN\n" + self.HORIZONTAL_LINE)
                print("LED CAN_RUN 测试通过\n关闭LED CAN_RUN\n" + self.HORIZONTAL_LINE)
                self.clearList(self.m_transmitData)
                self.m_transmitData[0] = 0x2f
                self.m_transmitData[1] = 0xf6
                self.m_transmitData[2] = 0x5f
                self.m_transmitData[3] = 0x03
                self.m_transmitData[4] = 0x00
                bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData)
                time.sleep(1)
                isLEDCANRunOK = True
            elif reply == QMessageBox.No:
                self.CAN_runLED = False
                self.pauseOption()
                if not self.is_running:
                    return False
                self.item_signal.emit([3, 2, 2, CANRunTest_time])
                # self.itemOperation(mTable, 3, 2, 2, CANRunTest_time)
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit("LED CAN_RUN 测试不通过\n" + self.HORIZONTAL_LINE)
                print("LED CAN_RUN 测试不通过\n" + self.HORIZONTAL_LINE)
                isLEDCANRunOK = False

            CANErrStart_time = time.time()
            self.pauseOption()
            if not self.is_running:
                return False
            self.item_signal.emit([4, 1, 0, ''])
            # self.itemOperation(mTable, 4, 1, 0, '')
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("2.进行 LED CAN_ERROR 测试\n\n")
            print("2.进行 LED CAN_ERROR 测试\n\n")
            self.clearList(self.m_transmitData)
            self.m_transmitData[0] = 0x2f
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x04
            self.m_transmitData[4] = 0x01
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData)
            CANErrEnd_time = time.time()
            CANErrTest_time = round(CANErrEnd_time - CANErrStart_time, 2)
            time.sleep(0.5)
            self.messageBox_signal.emit(['检测RUN &ERROR', 'CAN_ERROR指示灯是否点亮？'])
            reply = self.result_queue.get()
            # reply = QMessageBox.question(None, '检测CAN_RUN &CAN_ERROR', '指示灯是否点亮？',
            #                              QMessageBox.Yes | QMessageBox.No,
            #                              QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                self.CAN_errorLED = True
                self.pauseOption()
                if not self.is_running:
                    return False
                self.item_signal.emit([4, 2, 1, CANErrTest_time])
                # self.itemOperation(mTable, 4, 2, 1, CANErrTest_time)
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit("LED CAN_ERROR 测试通过\n关闭LED CAN_ERROR\n" + self.HORIZONTAL_LINE)
                print("LED CAN_ERROR 测试通过\n关闭LED CAN_ERROR\n" + self.HORIZONTAL_LINE)
                self.clearList(self.m_transmitData)
                self.m_transmitData[0] = 0x2f
                self.m_transmitData[1] = 0xf6
                self.m_transmitData[2] = 0x5f
                self.m_transmitData[3] = 0x04
                self.m_transmitData[4] = 0x00
                bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData)
                time.sleep(1)
                isLEDCANErrOK = True
            elif reply == QMessageBox.No:
                self.CAN_errorLED = False
                self.pauseOption()
                if not self.is_running:
                    return False
                self.item_signal.emit([4, 2, 2, CANErrTest_time])
                # self.itemOperation(mTable, 4, 2, 2, CANErrTest_time)
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit("LED CAN_ERROR 测试不通过\n" + self.HORIZONTAL_LINE)
                print("LED CAN_ERROR 测试不通过\n" + self.HORIZONTAL_LINE)
                isLEDCANErrOK = False
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("退出 LED TEST 模式\n" + self.HORIZONTAL_LINE)
            print("退出 LED TEST 模式\n" + self.HORIZONTAL_LINE)
            self.clearList(self.m_transmitData)
            self.m_transmitData[0] = 0x23
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x00
            self.m_transmitData[4] = 0x45
            self.m_transmitData[5] = 0x58
            self.m_transmitData[6] = 0x49
            self.m_transmitData[7] = 0x54
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData)

        else:
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("-----------进入 LED TEST 模式失败-----------\n\n")
            print("-----------进入 LED TEST 模式失败-----------\n\n")
            self.item_signal.emit([3, 2, 2, '进入模式失败'])
            self.item_signal.emit([4, 2, 2, '进入模式失败'])

        return True

    def normal_writeValuetoAO(self, value):
        bool_all = True
        for i in range(self.m_Channels):
            self.m_transmitData = [0x2b, 0x11, 0x64, i + 1, (value & 0xff), ((value >> 8) & 0xff), 0x00, 0x00]
            # self.m_transmitData[0] = 0x2b
            # self.m_transmitData[1] = 0x11
            # self.m_transmitData[2] = 0x64
            # self.m_transmitData[3] = i+1
            # self.m_transmitData[4] = (value & 0xff)
            # self.m_transmitData[5] = ((value >> 8) & 0xff)
            # self.m_transmitData[6] = 0x00
            # self.m_transmitData[7] = 0x00
            # print(f'{self.module_1}地址:{0x600+self.CANAddr_AO}')
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + self.CANAddr_AO), self.m_transmitData)
            bool_all = bool_all & bool_transmit
            # self.isPause()
            # if not self.isStop():
            #     return
        return bool_all

    def channelZero(self):
        self.pauseOption()
        if not self.is_running:
            return False
        self.result_signal.emit(f'对所有通道值进行归零处理' + self.HORIZONTAL_LINE)
        # print('1111111111111111')
        isZero = self.normal_writeValuetoAO(0)
        # print('2222222222222222')
        if isZero == True:
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f'所有通道归零成功' + self.HORIZONTAL_LINE)
            # self.isPause()
            # if not self.isStop():
            #     return

        else:
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f'所有通道归零失败' + self.HORIZONTAL_LINE)
            # self.isPause()
            # if not self.isStop():
            #     return
        return True

    # def generateExcel(self, station, module):
    #     book = xlwt.Workbook(encoding='utf-8')
    #     sheet = book.add_sheet('校准校验表', cell_overwrite_ok=True)
    #     # 如果出现报错：Exception: Attempt to overwrite cell: sheetname='sheet1' rowx=0 colx=0
    #     # 需要加上：cell_overwrite_ok=True)
    #     # 这是因为重复操作一个单元格导致的
    #     sheet.col(0).width = 256 * 12
    #     for i in range(99):
    #         #     sheet.w
    #         #     tall_style = xlwt.easyxf('font:height 240;')  # 36pt,类型小初的字号
    #         first_row = sheet.row(i)
    #         first_row.height_mismatch = True
    #         first_row.height = 20 * 20
    #
    #     # 为样式创建字体
    #     title_font = xlwt.Font()
    #     # 字体类型
    #     title_font.name = '宋'
    #     # 字体颜色
    #     title_font.colour_index = 0
    #     # 字体大小，11为字号，20为衡量单位
    #     title_font.height = 20 * 20
    #     # 字体加粗
    #     title_font.bold = True
    #
    #     # 设置单元格对齐方式
    #     title_alignment = xlwt.Alignment()
    #     # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    #     title_alignment.horz = 0x02
    #     # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    #     title_alignment.vert = 0x01
    #     # 设置自动换行
    #     title_alignment.wrap = 1
    #
    #     # 设置边框
    #     title_borders = xlwt.Borders()
    #     # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    #     # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    #     title_borders.left = 0
    #     title_borders.right = 0
    #     title_borders.top = 0
    #     title_borders.bottom = 0
    #
    #     # # 设置背景颜色
    #     # pattern = xlwt.Pattern()
    #     # # 设置背景颜色的模式
    #     # pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    #     # # 背景颜色
    #     # pattern.pattern_fore_colour = i
    #
    #     # # 初始化样式
    #     title_style = xlwt.XFStyle()
    #     title_style.borders = title_borders
    #     title_style.alignment = title_alignment
    #     title_style.font = title_font
    #     # # 设置文字模式
    #     # font.num_format_str = '#,##0.00'
    #
    #     # sheet.write(i, 0, u'字体', style0)
    #     # sheet.write(i, 1, u'背景', style1)
    #     # sheet.write(i, 2, u'对齐方式', style2)
    #     # sheet.write(i, 3, u'边框', style3)
    #
    #     # 合并单元格，合并第1行到第2行的第1列到第19列
    #     sheet.write_merge(0, 1, 0, 18, u'整机检验记录单V1.1', title_style)
    #
    #     # row3
    #     # 为样式创建字体
    #     row3_font = xlwt.Font()
    #     # 字体类型
    #     row3_font.name = '宋'
    #     # 字体颜色
    #     row3_font.colour_index = 0
    #     # 字体大小，11为字号，20为衡量单位
    #     row3_font.height = 10 * 20
    #     # 字体加粗
    #     row3_font.bold = False
    #
    #     # 设置单元格对齐方式
    #     row3_alignment = xlwt.Alignment()
    #     # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    #     row3_alignment.horz = 0x02
    #     # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    #     row3_alignment.vert = 0x01
    #     # 设置自动换行
    #     row3_alignment.wrap = 1
    #
    #     # 设置边框
    #     row3_borders = xlwt.Borders()
    #     # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    #     # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    #     row3_borders.left = 0
    #     row3_borders.right = 0
    #     row3_borders.top = 0
    #     row3_borders.bottom = 0
    #
    #     # # 设置背景颜色
    #     # pattern = xlwt.Pattern()
    #     # # 设置背景颜色的模式
    #     # pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    #     # # 背景颜色
    #     # pattern.pattern_fore_colour = i
    #
    #     # # 初始化样式
    #     row3_style = xlwt.XFStyle()
    #     row3_style.borders = row3_borders
    #     row3_style.alignment = row3_alignment
    #     row3_style.font = row3_font
    #     # # 设置文字模式
    #     # font.num_format_str = '#,##0.00'
    #
    #     # sheet.write(i, 0, u'字体', style0)
    #     # sheet.write(i, 1, u'背景', style1)
    #     # sheet.write(i, 2, u'对齐方式', style2)
    #     # sheet.write(i, 3, u'边框', style3)
    #
    #     sheet.write_merge(2, 2, 0, 2, 'PN：', row3_style)
    #     sheet.write_merge(2, 2, 3, 5, f'{self.module_pn}', row3_style)
    #     sheet.write_merge(2, 2, 6, 8, 'SN：', row3_style)
    #     sheet.write_merge(2, 2, 9, 11, f'{self.module_sn}', row3_style)
    #     sheet.write_merge(2, 2, 12, 14, 'REV：', row3_style)
    #     sheet.write_merge(2, 2, 15, 17, f'{self.module_rev}', row3_style)
    #
    #     # leftTitle
    #     # 为样式创建字体
    #     leftTitle_font = xlwt.Font()
    #     # 字体类型
    #     leftTitle_font.name = '宋'
    #     # 字体颜色
    #     leftTitle_font.colour_index = 0
    #     # 字体大小，11为字号，20为衡量单位
    #     leftTitle_font.height = 12 * 20
    #     # 字体加粗
    #     leftTitle_font.bold = True
    #
    #     # 设置单元格对齐方式
    #     leftTitle_alignment = xlwt.Alignment()
    #     # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    #     leftTitle_alignment.horz = 0x02
    #     # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    #     leftTitle_alignment.vert = 0x01
    #     # 设置自动换行
    #     leftTitle_alignment.wrap = 1
    #
    #     # 设置边框
    #     leftTitle_borders = xlwt.Borders()
    #     # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    #     # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    #     leftTitle_borders.left = 5
    #     leftTitle_borders.right = 5
    #     leftTitle_borders.top = 5
    #     leftTitle_borders.bottom = 5
    #
    #     # # 设置背景颜色
    #     # pattern = xlwt.Pattern()
    #     # # 设置背景颜色的模式
    #     # pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    #     # # 背景颜色
    #     # pattern.pattern_fore_colour = i
    #
    #     # # 初始化样式
    #     leftTitle_style = xlwt.XFStyle()
    #     leftTitle_style.borders = leftTitle_borders
    #     leftTitle_style.alignment = leftTitle_alignment
    #     leftTitle_style.font = leftTitle_font
    #     self.generalTest_row = 4
    #     sheet.write_merge(self.generalTest_row, self.generalTest_row + 1, 0, 0, '常规检测', leftTitle_style)
    #     self.CPU_row = 6
    #     sheet.write_merge(self.CPU_row, self.CPU_row + 7, 0, 0, 'CPU检测', leftTitle_style)
    #     self.DI_row = 14
    #     sheet.write_merge(self.DI_row, self.DI_row + 3, 0, 0, 'DI信号', leftTitle_style)
    #     self.DO_row = 18
    #     sheet.write_merge(self.DO_row, self.DO_row + 3, 0, 0, 'DO信号', leftTitle_style)
    #     self.AI_row = 22
    #
    #     sheet.write_merge(self.AI_row, self.AI_row + 1, 0, 0, 'AI信号', leftTitle_style)
    #     self.AO_row = self.AI_row + 2
    #
    #     sheet.write_merge(self.AO_row, self.AO_row + 1, 0, 0, 'AO信号', leftTitle_style)
    #     self.result_row = self.AO_row + 2
    #
    #     sheet.write_merge(self.result_row, self.result_row + 1, 0, 3, '整机检测结果', leftTitle_style)
    #
    #     # contentTitle
    #     # 为样式创建字体
    #     contentTitle_font = xlwt.Font()
    #     # 字体类型
    #     contentTitle_font.name = '宋'
    #     # 字体颜色
    #     contentTitle_font.colour_index = 0
    #     # 字体大小，11为字号，20为衡量单位
    #     contentTitle_font.height = 10 * 20
    #     # 字体加粗
    #     contentTitle_font.bold = False
    #
    #     # 设置单元格对齐方式
    #     contentTitle_alignment = xlwt.Alignment()
    #     # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    #     contentTitle_alignment.horz = 0x02
    #     # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    #     contentTitle_alignment.vert = 0x01
    #     # 设置自动换行
    #     contentTitle_alignment.wrap = 1
    #
    #     # 设置边框
    #     contentTitle_borders = xlwt.Borders()
    #     # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    #     # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    #     contentTitle_borders.left = 5
    #     contentTitle_borders.right = 5
    #     contentTitle_borders.top = 5
    #     contentTitle_borders.bottom = 5
    #
    #     # # 设置背景颜色
    #     # pattern = xlwt.Pattern()
    #     # # 设置背景颜色的模式
    #     # pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    #     # # 背景颜色
    #     # pattern.pattern_fore_colour = i
    #
    #     # # 初始化样式
    #     contentTitle_style = xlwt.XFStyle()
    #     contentTitle_style.borders = contentTitle_borders
    #     contentTitle_style.alignment = contentTitle_alignment
    #     contentTitle_style.font = contentTitle_font
    #
    #     sheet.write_merge(self.generalTest_row, self.generalTest_row, 1, 2, '外观', contentTitle_style)
    #     sheet.write(self.generalTest_row, 3, '---', contentTitle_style)
    #     sheet.write_merge(self.generalTest_row, self.generalTest_row, 4, 5, 'Run指示灯', contentTitle_style)
    #     sheet.write(self.generalTest_row, 6, '---', contentTitle_style)
    #     sheet.write_merge(self.generalTest_row, self.generalTest_row, 7, 8, 'Error指示灯', contentTitle_style)
    #     sheet.write(self.generalTest_row, 9, '---', contentTitle_style)
    #     sheet.write_merge(self.generalTest_row, self.generalTest_row, 10, 11, 'CAN_Run指示灯', contentTitle_style)
    #     sheet.write(self.generalTest_row, 12, '---', contentTitle_style)
    #     sheet.write_merge(self.generalTest_row, self.generalTest_row, 13, 14, 'CAN_Error指示灯', contentTitle_style)
    #     sheet.write(self.generalTest_row, 15, '---', contentTitle_style)
    #     sheet.write_merge(self.generalTest_row, self.generalTest_row, 16, 17, '拨码（预留）', contentTitle_style)
    #     sheet.write(self.generalTest_row, 18, '---', contentTitle_style)
    #     sheet.write_merge(self.generalTest_row + 1, self.generalTest_row + 1, 1, 2, '非测试项', contentTitle_style)
    #     sheet.write(self.generalTest_row + 1, 3, '---', contentTitle_style)
    #     sheet.write_merge(self.generalTest_row + 1, self.generalTest_row + 1, 4, 5, '------', contentTitle_style)
    #     sheet.write(self.generalTest_row + 1, 6, '---', contentTitle_style)
    #     sheet.write_merge(self.generalTest_row + 1, self.generalTest_row + 1, 7, 8, '------', contentTitle_style)
    #     sheet.write(self.generalTest_row + 1, 9, '---', contentTitle_style)
    #     sheet.write_merge(self.generalTest_row + 1, self.generalTest_row + 1, 10, 11, '------', contentTitle_style)
    #     sheet.write(self.generalTest_row + 1, 12, '---', contentTitle_style)
    #     sheet.write_merge(self.generalTest_row + 1, self.generalTest_row + 1, 13, 14, '------', contentTitle_style)
    #     sheet.write(self.generalTest_row + 1, 15, '---', contentTitle_style)
    #     sheet.write_merge(self.generalTest_row + 1, self.generalTest_row + 1, 16, 17, '------', contentTitle_style)
    #     sheet.write(self.generalTest_row + 1, 18, '---', contentTitle_style)
    #
    #     sheet.write_merge(self.CPU_row, self.CPU_row, 1, 2, '片外Flash读写', contentTitle_style)
    #     sheet.write(self.CPU_row, 3, '---', contentTitle_style)
    #     sheet.write_merge(self.CPU_row, self.CPU_row, 4, 5, 'MAC&序列号', contentTitle_style)
    #     sheet.write(self.CPU_row, self.CPU_row, '---', contentTitle_style)
    #     sheet.write_merge(self.CPU_row, self.CPU_row, 7, 8, '多功能按钮', contentTitle_style)
    #     sheet.write(self.CPU_row, 9, '---', contentTitle_style)
    #     sheet.write_merge(self.CPU_row, self.CPU_row, 10, 11, 'R/S拨杆', contentTitle_style)
    #     sheet.write(self.CPU_row, 12, '---', contentTitle_style)
    #     sheet.write_merge(self.CPU_row, self.CPU_row, 13, 14, '实时时钟', contentTitle_style)
    #     sheet.write(self.CPU_row, 15, '---', contentTitle_style)
    #     sheet.write_merge(self.CPU_row, self.CPU_row, 16, 17, 'SRAM', contentTitle_style)
    #     sheet.write(self.CPU_row, 18, '---', contentTitle_style)
    #     sheet.write_merge(self.CPU_row + 1, self.CPU_row + 1, 1, 2, '掉电保存', contentTitle_style)
    #     sheet.write(self.CPU_row + 1, 3, '---', contentTitle_style)
    #     sheet.write_merge(self.CPU_row + 1, self.CPU_row + 1, 4, 5, 'U盘', contentTitle_style)
    #     sheet.write(self.CPU_row + 1, 6, '---', contentTitle_style)
    #     sheet.write_merge(self.CPU_row + 1, self.CPU_row + 1, 7, 8, 'type-C', contentTitle_style)
    #     sheet.write(self.CPU_row + 1, 9, '---', contentTitle_style)
    #     sheet.write_merge(self.CPU_row + 1, self.CPU_row + 1, 10, 11, 'RS232通讯', contentTitle_style)
    #     sheet.write(self.CPU_row + 1, 12, '---', contentTitle_style)
    #     sheet.write_merge(self.CPU_row + 1, self.CPU_row + 1, 13, 14, 'RS485通讯', contentTitle_style)
    #     sheet.write(self.CPU_row + 1, 15, '---', contentTitle_style)
    #     sheet.write_merge(self.CPU_row + 1, self.CPU_row + 1, 16, 17, 'CAN通讯(预留)', contentTitle_style)
    #     sheet.write(self.CPU_row + 1, 18, '---', contentTitle_style)
    #
    #     sheet.write_merge(self.CPU_row + 2, self.CPU_row + 4, 1, 2, '输入通道', contentTitle_style)
    #     sheet.write(self.CPU_row + 2, 3, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 2, 4, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 2, 5, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 2, 6, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 2, 7, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 2, 8, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 2, 9, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 2, 10, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 2, 11, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 2, 12, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 2, 13, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 2, 14, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 2, 15, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 2, 16, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 2, 17, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 2, 18, '---', contentTitle_style)
    #
    #     sheet.write(self.CPU_row + 3, 3, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 3, 4, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 3, 5, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 3, 6, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 3, 7, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 3, 8, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 3, 9, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 3, 10, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 3, 11, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 3, 12, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 3, 13, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 3, 14, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 3, 15, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 3, 16, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 3, 17, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 3, 18, '---', contentTitle_style)
    #
    #     sheet.write(self.CPU_row + 4, 3, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 4, 4, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 4, 5, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 4, 6, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 4, 7, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 4, 8, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 4, 9, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 4, 10, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 4, 11, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 4, 12, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 4, 13, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 4, 14, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 4, 15, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 4, 16, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 4, 17, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 4, 18, '---', contentTitle_style)
    #
    #     sheet.write_merge(self.CPU_row + 5, self.CPU_row + 7, 1, 2, '输出通道', contentTitle_style)
    #     sheet.write(self.CPU_row + 5, 3, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 5, 4, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 5, 5, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 5, 6, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 5, 7, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 5, 8, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 5, 9, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 5, 10, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 5, 11, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 5, 12, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 5, 13, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 5, 14, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 5, 15, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 5, 16, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 5, 17, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 5, 18, '---', contentTitle_style)
    #
    #     sheet.write(self.CPU_row + 6, 3, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 6, 4, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 6, 5, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 6, 6, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 6, 7, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 6, 8, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 6, 9, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 6, 10, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 6, 11, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 6, 12, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 6, 13, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 6, 14, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 6, 15, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 6, 16, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 6, 17, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 6, 18, '---', contentTitle_style)
    #
    #     sheet.write(self.CPU_row + 7, 3, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 7, 4, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 7, 5, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 7, 6, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 7, 7, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 7, 8, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 7, 9, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 7, 10, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 7, 11, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 7, 12, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 7, 13, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 7, 14, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 7, 15, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 7, 16, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 7, 17, '---', contentTitle_style)
    #     sheet.write(self.CPU_row + 7, 18, '---', contentTitle_style)
    #     # DO
    #     sheet.write_merge(self.DO_row, self.DO_row, 1, 2, '通道号', contentTitle_style)
    #     sheet.write(self.DO_row, 3, 'CH1', contentTitle_style)
    #     sheet.write(self.DO_row, 4, 'CH2', contentTitle_style)
    #     sheet.write(self.DO_row, 5, 'CH3', contentTitle_style)
    #     sheet.write(self.DO_row, 6, 'CH4', contentTitle_style)
    #     sheet.write(self.DO_row, 7, 'CH5', contentTitle_style)
    #     sheet.write(self.DO_row, 8, 'CH6', contentTitle_style)
    #     sheet.write(self.DO_row, 9, 'CH7', contentTitle_style)
    #     sheet.write(self.DO_row, 10, 'CH8', contentTitle_style)
    #     sheet.write(self.DO_row, 11, 'CH9', contentTitle_style)
    #     sheet.write(self.DO_row, 12, 'CH10', contentTitle_style)
    #     sheet.write(self.DO_row, 13, 'CH11', contentTitle_style)
    #     sheet.write(self.DO_row, 14, 'CH12', contentTitle_style)
    #     sheet.write(self.DO_row, 15, 'CH13', contentTitle_style)
    #     sheet.write(self.DO_row, 16, 'CH14', contentTitle_style)
    #     sheet.write(self.DO_row, 17, 'CH15', contentTitle_style)
    #     sheet.write(self.DO_row, 18, 'CH16', contentTitle_style)
    #     sheet.write_merge(self.DO_row + 1, self.DO_row + 1, 1, 2, '是否合格', contentTitle_style)
    #     sheet.write(self.DO_row + 1, 3, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 1, 4, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 1, 5, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 1, 6, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 1, 7, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 1, 8, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 1, 9, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 1, 10, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 1, 11, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 1, 12, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 1, 13, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 1, 14, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 1, 15, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 1, 16, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 1, 17, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 1, 18, '---', contentTitle_style)
    #     sheet.write_merge(self.DO_row + 2, self.DO_row + 2, 1, 2, '通道号', contentTitle_style)
    #     sheet.write(self.DO_row + 2, 3, 'CH17', contentTitle_style)
    #     sheet.write(self.DO_row + 2, 4, 'CH18', contentTitle_style)
    #     sheet.write(self.DO_row + 2, 5, 'CH19', contentTitle_style)
    #     sheet.write(self.DO_row + 2, 6, 'CH20', contentTitle_style)
    #     sheet.write(self.DO_row + 2, 7, 'CH21', contentTitle_style)
    #     sheet.write(self.DO_row + 2, 8, 'CH22', contentTitle_style)
    #     sheet.write(self.DO_row + 2, 9, 'CH23', contentTitle_style)
    #     sheet.write(self.DO_row + 2, 10, 'CH24', contentTitle_style)
    #     sheet.write(self.DO_row + 2, 11, 'CH25', contentTitle_style)
    #     sheet.write(self.DO_row + 2, 12, 'CH26', contentTitle_style)
    #     sheet.write(self.DO_row + 2, 13, 'CH27', contentTitle_style)
    #     sheet.write(self.DO_row + 2, 14, 'CH28', contentTitle_style)
    #     sheet.write(self.DO_row + 2, 15, 'CH29', contentTitle_style)
    #     sheet.write(self.DO_row + 2, 16, 'CH30', contentTitle_style)
    #     sheet.write(self.DO_row + 2, 17, 'CH31', contentTitle_style)
    #     sheet.write(self.DO_row + 2, 18, 'CH32', contentTitle_style)
    #     sheet.write_merge(self.DO_row + 3, self.DO_row + 3, 1, 2, '是否合格', contentTitle_style)
    #     sheet.write(self.DO_row + 3, 3, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 3, 4, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 3, 5, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 3, 6, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 3, 7, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 3, 8, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 3, 9, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 3, 10, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 3, 11, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 3, 12, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 3, 13, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 3, 14, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 3, 15, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 3, 16, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 3, 17, '---', contentTitle_style)
    #     sheet.write(self.DO_row + 3, 18, '---', contentTitle_style)
    #
    #     # DI
    #     sheet.write_merge(self.DI_row, self.DI_row, 1, 2, '通道号', contentTitle_style)
    #     sheet.write(self.DI_row, 3, 'CH1', contentTitle_style)
    #     sheet.write(self.DI_row, 4, 'CH2', contentTitle_style)
    #     sheet.write(self.DI_row, 5, 'CH3', contentTitle_style)
    #     sheet.write(self.DI_row, 6, 'CH4', contentTitle_style)
    #     sheet.write(self.DI_row, 7, 'CH5', contentTitle_style)
    #     sheet.write(self.DI_row, 8, 'CH6', contentTitle_style)
    #     sheet.write(self.DI_row, 9, 'CH7', contentTitle_style)
    #     sheet.write(self.DI_row, 10, 'CH8', contentTitle_style)
    #     sheet.write(self.DI_row, 11, 'CH9', contentTitle_style)
    #     sheet.write(self.DI_row, 12, 'CH10', contentTitle_style)
    #     sheet.write(self.DI_row, 13, 'CH11', contentTitle_style)
    #     sheet.write(self.DI_row, 14, 'CH12', contentTitle_style)
    #     sheet.write(self.DI_row, 15, 'CH13', contentTitle_style)
    #     sheet.write(self.DI_row, 16, 'CH14', contentTitle_style)
    #     sheet.write(self.DI_row, 17, 'CH15', contentTitle_style)
    #     sheet.write(self.DI_row, 18, 'CH16', contentTitle_style)
    #     sheet.write_merge(self.DI_row + 1, self.DI_row + 1, 1, 2, '是否合格', contentTitle_style)
    #     sheet.write(self.DI_row + 1, 3, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 1, 4, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 1, 5, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 1, 6, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 1, 7, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 1, 8, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 1, 9, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 1, 10, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 1, 11, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 1, 12, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 1, 13, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 1, 14, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 1, 15, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 1, 16, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 1, 17, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 1, 18, '---', contentTitle_style)
    #     sheet.write_merge(self.DI_row + 2, self.DI_row + 2, 1, 2, '通道号', contentTitle_style)
    #     sheet.write(self.DI_row + 2, 3, 'CH17', contentTitle_style)
    #     sheet.write(self.DI_row + 2, 4, 'CH18', contentTitle_style)
    #     sheet.write(self.DI_row + 2, 5, 'CH19', contentTitle_style)
    #     sheet.write(self.DI_row + 2, 6, 'CH20', contentTitle_style)
    #     sheet.write(self.DI_row + 2, 7, 'CH21', contentTitle_style)
    #     sheet.write(self.DI_row + 2, 8, 'CH22', contentTitle_style)
    #     sheet.write(self.DI_row + 2, 9, 'CH23', contentTitle_style)
    #     sheet.write(self.DI_row + 2, 10, 'CH24', contentTitle_style)
    #     sheet.write(self.DI_row + 2, 11, 'CH25', contentTitle_style)
    #     sheet.write(self.DI_row + 2, 12, 'CH26', contentTitle_style)
    #     sheet.write(self.DI_row + 2, 13, 'CH27', contentTitle_style)
    #     sheet.write(self.DI_row + 2, 14, 'CH28', contentTitle_style)
    #     sheet.write(self.DI_row + 2, 15, 'CH29', contentTitle_style)
    #     sheet.write(self.DI_row + 2, 16, 'CH30', contentTitle_style)
    #     sheet.write(self.DI_row + 2, 17, 'CH31', contentTitle_style)
    #     sheet.write(self.DI_row + 2, 18, 'CH32', contentTitle_style)
    #     sheet.write_merge(self.DI_row + 3, self.DI_row + 3, 1, 2, '是否合格', contentTitle_style)
    #     sheet.write(self.DI_row + 3, 3, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 3, 4, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 3, 5, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 3, 6, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 3, 7, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 3, 8, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 3, 9, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 3, 10, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 3, 11, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 3, 12, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 3, 13, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 3, 14, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 3, 15, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 3, 16, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 3, 17, '---', contentTitle_style)
    #     sheet.write(self.DI_row + 3, 18, '---', contentTitle_style)
    #
    #     # AI
    #     sheet.write_merge(self.AI_row, self.AI_row + 1, 1, 1, '信号类型', contentTitle_style)
    #     sheet.write_merge(self.AI_row, self.AI_row + 1, 2, 3, '通道号', contentTitle_style)
    #     sheet.write_merge(self.AI_row, self.AI_row, 3 + 1, 5 + 1, '测试点1', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 3 + 1, '理论值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 4 + 1, '测试值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 5 + 1, '精度', contentTitle_style)
    #
    #     sheet.write_merge(self.AI_row, self.AI_row, 6 + 1, 8 + 1, '测试点2', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 6 + 1, '理论值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 7 + 1, '测试值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 8 + 1, '精度', contentTitle_style)
    #
    #     sheet.write_merge(self.AI_row, self.AI_row, 9 + 1, 11 + 1, '测试点3', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 9 + 1, '理论值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 10 + 1, '测试值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 11 + 1, '精度', contentTitle_style)
    #
    #     sheet.write_merge(self.AI_row, self.AI_row, 12 + 1, 14 + 1, '测试点4', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 12 + 1, '理论值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 13 + 1, '测试值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 14 + 1, '精度', contentTitle_style)
    #
    #     sheet.write_merge(self.AI_row, self.AI_row, 15 + 1, 17 + 1, '测试点5', contentTitle_style)
    #
    #     sheet.write(self.AI_row + 1, 15 + 1, '理论值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 16 + 1, '测试值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 17 + 1, '精度', contentTitle_style)
    #     # sheet.write(self.AI_row, 18, '', contentTitle_style)
    #     # sheet.write(self.AI_row + 1, 18, '', contentTitle_style)
    #
    #     # AO
    #     sheet.write_merge(self.AO_row, self.AO_row + 1, 1, 1, '信号类型', contentTitle_style)
    #     sheet.write_merge(self.AO_row, self.AO_row + 1, 2, 2 + 1, '通道号', contentTitle_style)
    #     sheet.write_merge(self.AO_row, self.AO_row, 3 + 1, 5 + 1, '测试点1', contentTitle_style)
    #     sheet.write(self.AO_row + 1, 3 + 1, '理论值', contentTitle_style)
    #     sheet.write(self.AO_row + 1, 4 + 1, '测试值', contentTitle_style)
    #     sheet.write(self.AO_row + 1, 5 + 1, '精度', contentTitle_style)
    #
    #     sheet.write_merge(self.AO_row, self.AO_row, 6 + 1, 8 + 1, '测试点2', contentTitle_style)
    #     sheet.write(self.AO_row + 1, 6 + 1, '理论值', contentTitle_style)
    #     sheet.write(self.AO_row + 1, 7 + 1, '测试值', contentTitle_style)
    #     sheet.write(self.AO_row + 1, 8 + 1, '精度', contentTitle_style)
    #
    #     sheet.write_merge(self.AO_row, self.AO_row, 9 + 1, 11 + 1, '测试点3', contentTitle_style)
    #     sheet.write(self.AO_row + 1, 9 + 1, '理论值', contentTitle_style)
    #     sheet.write(self.AO_row + 1, 10 + 1, '测试值', contentTitle_style)
    #     sheet.write(self.AO_row + 1, 11 + 1, '精度', contentTitle_style)
    #
    #     sheet.write_merge(self.AO_row, self.AO_row, 12 + 1, 14 + 1, '测试点4', contentTitle_style)
    #     sheet.write(self.AO_row + 1, 12 + 1, '理论值', contentTitle_style)
    #     sheet.write(self.AO_row + 1, 13 + 1, '测试值', contentTitle_style)
    #     sheet.write(self.AO_row + 1, 14 + 1, '精度', contentTitle_style)
    #
    #     sheet.write_merge(self.AO_row, self.AO_row, 15 + 1, 17 + 1, '测试点5', contentTitle_style)
    #     sheet.write(self.AO_row + 1, 15 + 1, '理论值', contentTitle_style)
    #     sheet.write(self.AO_row + 1, 16 + 1, '测试值', contentTitle_style)
    #     sheet.write(self.AO_row + 1, 17 + 1, '精度', contentTitle_style)
    #     # sheet.write(self.AO_row, 18, '', contentTitle_style)
    #     # sheet.write(self.AO_row + 1, 18, '', contentTitle_style)
    #
    #     # 结果
    #     sheet.write_merge(self.result_row, self.result_row, 4, 5, '□ 合格', contentTitle_style)
    #     sheet.write_merge(self.result_row, self.result_row, 6, 18, ' ', contentTitle_style)
    #     sheet.write_merge(self.result_row + 1, self.result_row + 1, 4, 5, '□ 不合格', contentTitle_style)
    #     sheet.write_merge(self.result_row + 1, self.result_row + 1, 6, 18, ' ', contentTitle_style)
    #
    #     # 补充说明
    #     sheet.write(self.result_row + 2, 0, '补充说明：', contentTitle_style)
    #     sheet.write_merge(self.result_row + 2, self.result_row + 2, 1, 18,
    #                       'AI/AO信号检验要记录数据，电压和电流的精度为1‰以下为合格；其他测试项合格打“√”，否则打“×”',
    #                       contentTitle_style)
    #
    #     # 检测信息
    #     sheet.write_merge(self.result_row + 3, self.result_row + 3, 0, 1, '检验员：', contentTitle_style)
    #     sheet.write_merge(self.result_row + 3, self.result_row + 3, 2, 3, '555', contentTitle_style)
    #     sheet.write_merge(self.result_row + 3, self.result_row + 3, 4, 5, '检验日期：', contentTitle_style)
    #     sheet.write_merge(self.result_row + 3, self.result_row + 3, 6, 8, f'{time.strftime("%Y-%m-%d %H：%M：%S")}',
    #                       contentTitle_style)
    #     sheet.write_merge(self.result_row + 3, self.result_row + 3, 9, 10, '审核：', contentTitle_style)
    #     sheet.write_merge(self.result_row + 3, self.result_row + 3, 11, 13, ' ', contentTitle_style)
    #     sheet.write_merge(self.result_row + 3, self.result_row + 3, 14, 15, '审核日期：', contentTitle_style)
    #     sheet.write_merge(self.result_row + 3, self.result_row + 3, 16, 18, ' ', contentTitle_style)
    #
    #
    #     if module == 'DI':
    #         self.fillInDIData(station, book, sheet)
    #     elif module == 'DO':
    #         self.fillInDOData(station, book, sheet)
    #     # elif module == 'AI':
    #     #     # print('打印AI检测结果')
    #     #     self.fillInAIData(station, book, sheet)
    #     # elif module == 'AO':
    #     #     # print('打印AI检测结果')
    #     #     self.fillInAOData(station, book, sheet)

    def fillInDIData(self, station, book, sheet):
        self.generalTest_row = 4
        # 通过单元格样式
        # 为样式创建字体
        pass_font = xlwt.Font()
        # 字体类型
        pass_font.name = '宋'
        # 字体颜色
        pass_font.colour_index = 0
        # 字体大小，11为字号，20为衡量单位
        pass_font.height = 10 * 20
        # 字体加粗
        pass_font.bold = False

        # 设置单元格对齐方式
        pass_alignment = xlwt.Alignment()
        # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        pass_alignment.horz = 0x02
        # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        pass_alignment.vert = 0x01
        # 设置自动换行
        pass_alignment.wrap = 1

        # 设置边框
        pass_borders = xlwt.Borders()
        # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
        # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
        pass_borders.left = 5
        pass_borders.right = 5
        pass_borders.top = 5
        pass_borders.bottom = 5

        # 设置背景颜色
        pass_pattern = xlwt.Pattern()
        # 设置背景颜色的模式
        pass_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        # 背景颜色
        pass_pattern.pattern_fore_colour = 3

        # # 初始化样式
        pass_style = xlwt.XFStyle()
        pass_style.borders = pass_borders
        pass_style.alignment = pass_alignment
        pass_style.font = pass_font
        pass_style.pattern = pass_pattern

        # 未通过单元格样式
        # 为样式创建字体
        fail_font = xlwt.Font()
        # 字体类型
        fail_font.name = '宋'
        # 字体颜色
        fail_font.colour_index = 1
        # 字体大小，11为字号，20为衡量单位
        fail_font.height = 10 * 20
        # 字体加粗
        fail_font.bold = True

        # 设置单元格对齐方式
        fail_alignment = xlwt.Alignment()
        # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        fail_alignment.horz = 0x02
        # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        fail_alignment.vert = 0x01
        # 设置自动换行
        fail_alignment.wrap = 1

        # 设置边框
        fail_borders = xlwt.Borders()
        # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
        # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
        fail_borders.left = 5
        fail_borders.right = 5
        fail_borders.top = 5
        fail_borders.bottom = 5

        # 设置背景颜色
        fail_pattern = xlwt.Pattern()
        # 设置背景颜色的模式
        fail_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        # 背景颜色
        fail_pattern.pattern_fore_colour = 2

        # # 初始化样式
        fail_style = xlwt.XFStyle()
        fail_style.borders = fail_borders
        fail_style.alignment = fail_alignment
        fail_style.font = fail_font
        fail_style.pattern = fail_pattern

        # 提示警告单元格样式
        # 为样式创建字体
        warning_font = xlwt.Font()
        # 字体类型
        warning_font.name = '宋'
        # 字体颜色
        warning_font.colour_index = 2
        # 字体大小，11为字号，20为衡量单位
        warning_font.height = 10 * 20
        # 字体加粗
        warning_font.bold = True

        # 设置单元格对齐方式
        warning_alignment = xlwt.Alignment()
        # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        warning_alignment.horz = 0x02
        # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        warning_alignment.vert = 0x01
        # 设置自动换行
        warning_alignment.wrap = 1

        # 设置边框
        warning_borders = xlwt.Borders()
        # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
        # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
        warning_borders.left = 5
        warning_borders.right = 5
        warning_borders.top = 5
        warning_borders.bottom = 5

        # 设置背景颜色
        warning_pattern = xlwt.Pattern()
        # 设置背景颜色的模式
        warning_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        # 背景颜色
        warning_pattern.pattern_fore_colour = 5

        # # 初始化样式
        warning_style = xlwt.XFStyle()
        warning_style.borders = warning_borders
        warning_style.alignment = warning_alignment
        warning_style.font = warning_font
        warning_style.pattern = warning_pattern

        if station and self.testNum == 0:
            name_save = '合格'
        if station and self.testNum != 0:
            name_save = '部分合格'
        elif not station:
            name_save = '不合格'

        if self.appearance:
            sheet.write(self.generalTest_row, 3, '√', pass_style)
        elif not self.appearance:
            sheet.write(self.generalTest_row, 3, '×', fail_style)
            self.errorNum += 1
            self.errorInf += f'\n{self.errorNum})外观存在瑕疵 '
        if not self.isTestRunErr:
            sheet.write(self.generalTest_row, 6, '未检测', warning_style)
            sheet.write(self.generalTest_row, 9, '未检测', warning_style)
        if not self.isTestCANRunErr:
            sheet.write(self.generalTest_row, 12, '未检测', warning_style)
            sheet.write(self.generalTest_row, 15, '未检测', warning_style)
        if self.CAN_runLED and self.isTestCANRunErr:
            sheet.write(self.generalTest_row, 12, '√', pass_style)
        elif not self.CAN_runLED and self.isTestCANRunErr:
            sheet.write(self.generalTest_row, 12, '×', fail_style)
            self.errorNum += 1
            self.errorInf += f'\n{self.errorNum})CAN_RUN指示灯未亮 '
        if self.CAN_errorLED and self.isTestCANRunErr:
            sheet.write(self.generalTest_row, 15, '√', pass_style)
        elif not self.CAN_errorLED and self.isTestCANRunErr:
            sheet.write(self.generalTest_row, 15, '×', fail_style)
            self.errorNum += 1
            self.errorInf += f'\n{self.errorNum})CAN_ERROR指示灯未亮 '
        if self.runLED and self.isTestRunErr:
            sheet.write(self.generalTest_row, 6, '√', pass_style)
        elif not self.runLED and self.isTestRunErr:
            sheet.write(self.generalTest_row, 6, '×', fail_style)
            self.errorNum += 1
            self.errorInf += f'\n{self.errorNum})RUN指示灯未亮 '
        if self.errorLED and self.isTestRunErr:
            sheet.write(self.generalTest_row, 9, '√', pass_style)
        elif not self.errorLED and self.isTestRunErr:
            sheet.write(self.generalTest_row, 9, '×', fail_style)
            self.errorNum += 1
            self.errorInf += f'\n{self.errorNum})ERROE指示灯未亮 '
        # #填写信号类型、通道号、测试点数据
        #     if self.isAITestVol:
        #         all_row = 9 + 4 + 4 + (2 + self.AI_Channels) + 2  # CPU + DI + DO + AI + AO
        #         sheet.write_merge(self.AI_row + 2,self.AI_row + 1 + self.AI_Channels,1,1,'电压',pass_style)
        #         for i in range(self.AI_Channels):
        #             #通道号
        #             sheet.write(self.AI_row + 2 + i, 2, f'CH{i+1}', pass_style)
        #             for i in range(5):
        #                 #理论值
        #                 sheet.write(self.AI_row + 2 + i, 3 + 3 * i, f'{self.voltageTheory[i]}', pass_style)
        #                 #测试值
        #                 sheet.write(self.AI_row + 2 + i, 4 + 3 * i, f'{self.volReceValue[i]}', pass_style)
        #                 # 精度
        #                 if abs(self.volPrecision[i]) < 2:
        #                     sheet.write(self.AI_row + 2 + i, 5 + 3 * i, f'{self.volPrecision[i]}‰', pass_style)
        #                 else:
        #                     sheet.write(self.AI_row + 2 + i, 5 + 3 * i, f'{self.volPrecision[i]}‰', fail_style)
        #     if self.isAITestVol and self.isAITestCur:
        #         all_row = 9 + 4 + 4 + (2 + 2 * self.AI_Channels) + 2  # CPU + DI + DO + AI + AO
        #         sheet.write_merge(self.AI_row + 2 + self.AI_Channels,self.AI_row + 1 + 2 * self.AI_Channels,1,1,'电流',pass_style)
        #         for i in range(self.AI_Channels):
        #             #通道号
        #             sheet.write(self.AI_row + 6 + i, 2, f'CH{i+1}', pass_style)
        #             for i in range(5):
        #                 #理论值
        #                 sheet.write(self.AI_row + 6 + i, 3 + 3 * i, f'{self.currentTheory[i]}', pass_style)
        #                 #测试值
        #                 sheet.write(self.AI_row + 6 + i, 4 + 3 * i, f'{self.curReceValue[i]}', pass_style)
        #                 # 精度
        #                 if abs(self.curPrecision[i]) < 2:
        #                     sheet.write(self.AI_row + 6 + i, 5 + 3 * i, f'{self.curPrecision[i]}‰', pass_style)
        #                 else:
        #                     sheet.write(self.AI_row + 6 + i, 5 + 3 * i, f'{self.curPrecision[i]}‰', fail_style)
        #     if not self.isAITestVol and self.isAITestCur:
        #         all_row = 9 + 4 + 4 + (2 + self.AI_Channels) + 2  # CPU + DI + DO + AI + AO
        #         sheet.write_merge(self.AI_row + 2, self.AI_row + 1 + self.AI_Channels, 1, 1, '电流', pass_style)
        #         for i in range(self.AI_Channels):
        #             # 通道号
        #             sheet.write(self.AI_row + 2 + i, 2, f'CH{i + 1}', pass_style)
        #             for i in range(5):
        #                 # 理论值
        #                 sheet.write(self.AI_row + 2 + i, 3 + 3 * i, f'{self.currentTheory[i]}', pass_style)
        #                 # 测试值
        #                 sheet.write(self.AI_row + 2 + i, 4 + 3 * i, f'{self.curReceValue[i]}', pass_style)
        #                 # 精度
        #                 if abs(self.curPrecision[i]) < 2:
        #                     sheet.write(self.AI_row + 2 + i, 5 + 3 * i, f'{self.curPrecision[i]}‰', pass_style)
        #                 else:
        #                     sheet.write(self.AI_row + 2 + i, 5 + 3 * i, f'{self.curPrecision[i]}‰', fail_style)
        #     if not self.isAITestVol and not self.isAITestCur:
        #         all_row = 9 + 4 + 4 + 2 + 2  # CPU + DI + DO + AI + AO
        # 填写通道状态

        for i in range(self.m_Channels):
            if i < 16:
                if self.DIDataCheck[i]:
                    sheet.write(self.DI_row + 1, 3 + i, '√', pass_style)
                elif not self.DIDataCheck[i]:
                    sheet.write(self.DI_row + 1, 3 + i, '×', fail_style)
                    self.errorNum += 1
                    self.errorInf += f'\n{self.errorNum})通道{i + 1} 灯未亮 '
            else:
                if self.DIDataCheck[i]:
                    sheet.write(self.DI_row + 3, i - 13, '√', pass_style)
                elif not self.DIDataCheck[i]:
                    sheet.write(self.DI_row + 3, i - 13, '×', fail_style)
                    self.errorNum += 1
                    self.errorInf += f'\n{self.errorNum})通道{i + 1} 灯未亮 '

        self.isDIPassTest = (((((self.isDIPassTest & self.isLEDRunOK) & self.isLEDErrOK) &
                               self.CAN_runLED) & self.CAN_errorLED) & self.appearance)
        # self.showInf(f'self.isLEDRunOK:{self.isLEDRunOK}')
        all_row = 9 + 4 + 4 + 2 + 2  # CPU + DI + DO + AI + AO
        if self.isDIPassTest and self.testNum == 0:
            name_save = '合格'
            sheet.write(self.generalTest_row + all_row + 1, 4, '■ 合格', pass_style)
            self.label_signal.emit(['pass', '全部通过'])
            self.print_signal.emit(
                [f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}', 'PASS', ''])
            # self.label.setStyleSheet(self.testState_qss['pass'])
            # self.label.setText('通过')
        if self.isDIPassTest and self.testNum > 0:
            name_save = '部分合格'
            sheet.write(self.generalTest_row + all_row + 1, 4, '■ 部分合格', pass_style)
            sheet.write(self.generalTest_row + all_row + 1, 6,
                        '------------------ 注意：有部分项目未测试！！！ ------------------', warning_style)
            self.label_signal.emit(['testing', '部分通过'])
            self.print_signal.emit(
                [f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}', 'PASS', ''])
            # self.label.setStyleSheet(self.testState_qss['testing'])
            # self.label.setText('部分通过')
        elif not self.isDIPassTest:
            if not self.isPassOdd:
                self.errorNum += 1
                self.errorInf += f'\n{self.errorNum})奇数指示灯全亮时有问题 '
            if not self.isPassEven:
                self.errorNum += 1
                self.errorInf += f'\n{self.errorNum})偶数指示灯全亮时有问题 '
            name_save = '不合格'
            sheet.write(self.generalTest_row + all_row + 2, 4, '■ 不合格', fail_style)
            sheet.write(self.generalTest_row + all_row + 2, 6, f'不合格原因：{self.errorInf}', warning_style)
            self.label_signal.emit(['fail', '未通过'])
            self.print_signal.emit(
                [f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}', 'FAIL', self.errorInf])
            # self.label.setStyleSheet(self.testState_qss['fail'])
            # self.label.setText('未通过')
        # book.save(self.saveDir + f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}.xls')
        self.saveExcel_signal.emit([book, f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}.xls'])

    # @abstractmethod
    def fillInDOData(self, station, book, sheet):
        # 通过单元格样式
        # 为样式创建字体
        pass_font = xlwt.Font()
        # 字体类型
        pass_font.name = '宋'
        # 字体颜色
        pass_font.colour_index = 0
        # 字体大小，11为字号，20为衡量单位
        pass_font.height = 10 * 20
        # 字体加粗
        pass_font.bold = False

        # 设置单元格对齐方式
        pass_alignment = xlwt.Alignment()
        # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        pass_alignment.horz = 0x02
        # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        pass_alignment.vert = 0x01
        # 设置自动换行
        pass_alignment.wrap = 1

        # 设置边框
        pass_borders = xlwt.Borders()
        # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
        # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
        pass_borders.left = 5
        pass_borders.right = 5
        pass_borders.top = 5
        pass_borders.bottom = 5

        # 设置背景颜色
        pass_pattern = xlwt.Pattern()
        # 设置背景颜色的模式
        pass_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        # 背景颜色
        pass_pattern.pattern_fore_colour = 3

        # # 初始化样式
        pass_style = xlwt.XFStyle()
        pass_style.borders = pass_borders
        pass_style.alignment = pass_alignment
        pass_style.font = pass_font
        pass_style.pattern = pass_pattern

        # 未通过单元格样式
        # 为样式创建字体
        fail_font = xlwt.Font()
        # 字体类型
        fail_font.name = '宋'
        # 字体颜色
        fail_font.colour_index = 1
        # 字体大小，11为字号，20为衡量单位
        fail_font.height = 10 * 20
        # 字体加粗
        fail_font.bold = True

        # 设置单元格对齐方式
        fail_alignment = xlwt.Alignment()
        # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        fail_alignment.horz = 0x02
        # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        fail_alignment.vert = 0x01
        # 设置自动换行
        fail_alignment.wrap = 1

        # 设置边框
        fail_borders = xlwt.Borders()
        # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
        # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
        fail_borders.left = 5
        fail_borders.right = 5
        fail_borders.top = 5
        fail_borders.bottom = 5

        # 设置背景颜色
        fail_pattern = xlwt.Pattern()
        # 设置背景颜色的模式
        fail_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        # 背景颜色
        fail_pattern.pattern_fore_colour = 2

        # # 初始化样式
        fail_style = xlwt.XFStyle()
        fail_style.borders = fail_borders
        fail_style.alignment = fail_alignment
        fail_style.font = fail_font
        fail_style.pattern = fail_pattern

        # 提示警告单元格样式
        # 为样式创建字体
        warning_font = xlwt.Font()
        # 字体类型
        warning_font.name = '宋'
        # 字体颜色
        warning_font.colour_index = 2
        # 字体大小，11为字号，20为衡量单位
        warning_font.height = 10 * 20
        # 字体加粗
        warning_font.bold = True

        # 设置单元格对齐方式
        warning_alignment = xlwt.Alignment()
        # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        warning_alignment.horz = 0x02
        # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        warning_alignment.vert = 0x01
        # 设置自动换行
        warning_alignment.wrap = 1

        # 设置边框
        warning_borders = xlwt.Borders()
        # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
        # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
        warning_borders.left = 5
        warning_borders.right = 5
        warning_borders.top = 5
        warning_borders.bottom = 5

        # 设置背景颜色
        warning_pattern = xlwt.Pattern()
        # 设置背景颜色的模式
        warning_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        # 背景颜色
        warning_pattern.pattern_fore_colour = 5

        # # 初始化样式
        warning_style = xlwt.XFStyle()
        warning_style.borders = warning_borders
        warning_style.alignment = warning_alignment
        warning_style.font = warning_font
        warning_style.pattern = warning_pattern

        # if station and self.testNum == 0:
        #     name_save = '合格'
        # if station and self.testNum != 0:
        #     name_save = '部分合格'
        # elif not station:
        #     name_save = '不合格'

        if self.appearance:
            sheet.write(self.generalTest_row, 3, '√', pass_style)
        elif not self.appearance:
            sheet.write(self.generalTest_row, 3, '×', fail_style)
            self.errorNum += 1
            self.errorInf += f'\n{self.errorNum})外观存在瑕疵 '
        if not self.isTestRunErr:
            sheet.write(self.generalTest_row, 6, '未检测', warning_style)
            sheet.write(self.generalTest_row, 9, '未检测', warning_style)
        if not self.isTestCANRunErr:
            sheet.write(self.generalTest_row, 12, '未检测', warning_style)
            sheet.write(self.generalTest_row, 15, '未检测', warning_style)
        if self.CAN_runLED and self.isTestCANRunErr:
            sheet.write(self.generalTest_row, 12, '√', pass_style)
        elif not self.CAN_runLED and self.isTestCANRunErr:
            sheet.write(self.generalTest_row, 12, '×', fail_style)
            self.errorNum += 1
            self.errorInf += f'\n{self.errorNum})CAN_RUN指示灯未亮 '
        if self.CAN_errorLED and self.isTestCANRunErr:
            sheet.write(self.generalTest_row, 15, '√', pass_style)
        elif not self.CAN_errorLED and self.isTestCANRunErr:
            sheet.write(self.generalTest_row, 15, '×', fail_style)
            self.errorNum += 1
            self.errorInf += f'\n{self.errorNum})CAN_ERROR指示灯未亮 '
        if self.runLED and self.isTestRunErr:
            sheet.write(self.generalTest_row, 6, '√', pass_style)
        elif not self.runLED and self.isTestRunErr:
            sheet.write(self.generalTest_row, 6, '×', fail_style)
            self.errorNum += 1
            self.errorInf += f'\n{self.errorNum})RUN指示灯未亮 '
        if self.errorLED and self.isTestRunErr:
            sheet.write(self.generalTest_row, 9, '√', pass_style)
        elif not self.errorLED and self.isTestRunErr:
            sheet.write(self.generalTest_row, 9, '×', fail_style)
            self.errorNum += 1
            self.errorInf += f'\n{self.errorNum})ERROE指示灯未亮 '

        # 填写通道状态
        for i in range(self.m_Channels):
            if i < 16:
                if self.DODataCheck[i]:
                    sheet.write(self.DO_row + 1, 3 + i, '√', pass_style)
                elif not self.DODataCheck[i]:
                    sheet.write(self.DO_row + 1, 3 + i, '×', fail_style)
                    self.errorNum += 1
                    self.errorInf += f'\n{self.errorNum}.通道{i + 1}灯未亮 '
            else:
                if self.DODataCheck[i]:
                    sheet.write(self.DO_row + 3, i - 13, '√', pass_style)
                elif not self.DODataCheck[i]:
                    sheet.write(self.DO_row + 3, i - 13, '×', fail_style)
                    self.errorNum += 1
                    self.errorInf += f'\n{self.errorNum}.通道{i + 1}灯未亮 '

        self.isDOPassTest = (((((self.isDOPassTest & self.isLEDRunOK) & self.isLEDErrOK) &
                               self.CAN_runLED) & self.CAN_errorLED) & self.appearance)
        # self.showInf(f'self.isLEDRunOK:{self.isLEDRunOK}')
        all_row = 9 + 4 + 4 + 2 + 2
        if self.isDOPassTest and self.testNum == 0:
            name_save = '合格'
            sheet.write(self.generalTest_row + all_row + 1, 4, '■ 合格', pass_style)
            self.label_signal.emit(['pass', '全部通过'])
            self.print_signal.emit([f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}', 'PASS', ''])
            # self.label.setStyleSheet(self.testState_qss['pass'])
            # self.label.setText('通过')
        if self.isDOPassTest and self.testNum > 0:
            name_save = '部分合格'
            sheet.write(self.generalTest_row + all_row + 1, 4, '■ 部分合格', pass_style)
            sheet.write(self.generalTest_row + all_row + 1, 6,
                        '------------------ 注意：有部分项目未测试！！！ ------------------', warning_style)
            self.label_signal.emit(['testing', '部分通过'])
            self.print_signal.emit([f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}', 'PASS', ''])
            # self.label.setStyleSheet(self.testState_qss['testing'])
            # self.label.setText('部分通过')
        elif not self.isDOPassTest:
            if not self.isPassOdd:
                self.errorNum += 1
                self.errorInf += f'\n{self.errorNum})奇数指示灯全亮时有问题 '
            if not self.isPassEven:
                self.errorNum += 1
                self.errorInf += f'\n{self.errorNum})偶数指示灯全亮时有问题 '
            name_save = '不合格'
            sheet.write(self.generalTest_row + all_row + 2, 4, '■ 不合格', fail_style)
            sheet.write(self.generalTest_row + all_row + 2, 6, f'不合格原因：{self.errorInf}', fail_style)
            self.label_signal.emit(['fail', '未通过'])
            self.print_signal.emit(
                [f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}', 'FAIL', self.errorInf])
            # self.label.setStyleSheet(self.testState_qss['fail'])
            # self.label.setText('未通过')
        # book.save(self.saveDir + f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}.xls')
        self.saveExcel_signal.emit([book, f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}.xls'])

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
                    break

    def stop_work(self):
        self.resume_work()
        self.is_running = False

