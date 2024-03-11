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
    messageBox_signal = pyqtSignal(list)
    pic_messageBox_signal = pyqtSignal(list)
    messageBox_oneButton_signal = pyqtSignal(list)
    gif_messageBox_signal = pyqtSignal(list)
    checkBox_messageBox_signal = pyqtSignal(list)
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
                       [0x00, 0x00, 0x00, 0x80], [0xff, 0xff, 0xff, 0xff], [0x00, 0x00, 0x00, 0x00]
                       ]
    # ,
    # [0xaa, 0xaa, 0xaa, 0xaa], [0x55, 0x55, 0x55, 0x55]

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

    def __init__(self, inf_DIDOlist: list, result_queue, appearance, testFlage):
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
        # #print(f'inf_AOlist[5] = {inf_AOlist[5]}')
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
        self.led_channel = {0: '00', 1: '01', 2: '02', 3: '03', 4: '04', 5: '05', 6: '06', 7: '07',
                            8: '10', 9: '11', 10: '12', 11: '13', 12: '14', 13: '15', 14: '16', 15: '17'}
        self.current_dir = os.getcwd().replace('\\', '/') + "/_internal"

        self.stopChannel = 2

        self.channelLED = [True,True,True,True,
                         True,True,True,True,
                         True,True,True,True,
                         True,True,True,True]
    def DIDOOption(self):
        isExcel = True
        if self.testFlage == 'DI':
            if not self.testDI():
                self.result_signal.emit("测试停止,后续测试全部取消" + self.HORIZONTAL_LINE)
                isExcel = False
        elif self.testFlage == 'DO':
            if not self.testDO():
                self.result_signal.emit("测试停止,后续测试全部取消" + self.HORIZONTAL_LINE)
                isExcel = False

        self.pauseOption()
        if not self.is_running:
            isExcel = False
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
            self.result_signal.emit('校准校验表生成失败…………' + self.HORIZONTAL_LINE)

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
            # print("--------------进入 LED TEST模式-------------")
            # self.channelZero()
            self.m_transmitData[0] = 0x23
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x00
            self.m_transmitData[4] = 0x53
            self.m_transmitData[5] = 0x54
            self.m_transmitData[6] = 0x41
            self.m_transmitData[7] = 0x52
            isLEDTest, whatEver = CAN_option.transmitCAN((0x600 + int(self.CANAddr_DI)), self.m_transmitData, 1)
            if isLEDTest:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit("-----------进入 LED TEST 模式成功-----------\n")
                # print("-----------进入 LED TEST 模式成功-----------" + self.HORIZONTAL_LINE)

                if self.isTestRunErr:
                    if not self.testRunErr(int(self.CANAddr_DI), 'DI'):
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
                                                                       self.m_transmitData, 1)
                if bool_transmit:
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    self.result_signal.emit("成功退出 LED TEST 模式。\n" + self.HORIZONTAL_LINE)
                    # print("成功退出 LED TEST 模式\n" + self.HORIZONTAL_LINE)
                else:
                    self.result_signal.emit("退出 LED TEST 模式失败！\n" + self.HORIZONTAL_LINE)
            else:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit("-----------进入 LED TEST 模式失败-----------\n\n")
                # print("-----------进入 LED TEST 模式失败-----------\n\n")
                self.item_signal.emit([0, 2, 2, '进入模式失败'])
                self.item_signal.emit([1, 2, 2, '进入模式失败'])

        self.item_signal.emit([4, 1, 0, ''])
        self.DIDataCheck = [True for i in range(32)]
        for i in range(self.loop_num):
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f"第{i + 1}次测试" + self.HORIZONTAL_LINE)
            CAN_option.Can_DLL.VCI_ClearBuffer(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX, CAN_option.CAN_INDEX)
            time.sleep(0.2)
            self.pauseOption()
            if not self.is_running:
                return False

            # 奇数通道检测
            self.pauseOption()
            if not self.is_running:
                return False
            # self.messageBox_oneButton_signal.emit(['操作提示','请观察待测模块通道指示灯亮灭情况。'])
            # reply = self.result_queue.get()
            # if reply == QMessageBox.AcceptRole:
            import threading
            odd_thread = threading.Thread(target=self.oddChannelTest,args=('DI',))
            odd_thread.start()
            image_ET1600_LED = self.current_dir + '/ET1600_odd.gif'
            self.gif_messageBox_signal.emit(
                [f'{self.module_type}通道检测', f'{self.module_type}通道指示灯是否如图所示闪烁？', image_ET1600_LED])
            reply = self.result_queue.get()

            self.stopChannel = reply
            if reply == QMessageBox.AcceptRole:
                self.isDIPassTest &= True
            else:
                self.isDIPassTest &= False

                self.checkBox_messageBox_signal.emit(['故障通道选择','请选择存在故障的通道灯。','odd'])
                messageList = self.result_queue.get()
                reply = messageList[0]
                channelLED_odd = messageList[1]
                if reply == QMessageBox.AcceptRole:
                    for ec in range(1,16,2):
                        self.channelLED[ec] &= channelLED_odd[ec]
                        if not channelLED_odd[ec]:
                            if ec < 8:
                                self.result_signal.emit('通道0' + str(ec)+'的通道灯存在问题。\n\n')
                            else:
                                self.result_signal.emit('通道' + str(ec+2) + '的通道灯存在问题。\n\n')
            # 等待子线程执行结束
            odd_thread.join()
            self.testByIndex(33, 'DI', 0)



            #偶数通道检测
            self.stopChannel = 2
            self.pauseOption()
            if not self.is_running:
                return False
            # self.messageBox_oneButton_signal.emit(['操作提示', '请观察待测模块通道指示灯亮灭情况。'])
            # reply = self.result_queue.get()
            # if reply == QMessageBox.AcceptRole:
            even_thread = threading.Thread(target=self.evenChannelTest,args=('DI',))
            even_thread.start()
            image_ET1600_LED = self.current_dir + '/ET1600_even.gif'
            self.gif_messageBox_signal.emit(
                [f'{self.module_type}通道检测', f'{self.module_type}通道指示灯是否如图所示闪烁？', image_ET1600_LED])
            reply = self.result_queue.get()

            self.stopChannel = reply
            if reply == QMessageBox.AcceptRole:
                self.isDIPassTest &= True
            else:
                self.isDIPassTest &= False
                self.checkBox_messageBox_signal.emit(['故障通道选择', '请选择存在故障的通道灯。','even'])
                messageList = self.result_queue.get()
                reply = messageList[0]
                channelLED_even = messageList[1]
                if reply == QMessageBox.AcceptRole:
                    for ec in range(0, 15, 2):
                        self.channelLED[ec] &= channelLED_even[ec]
                        if not channelLED_even[ec]:
                            if ec < 7:
                                self.result_signal.emit('通道0' + str(ec) + '的通道灯存在问题。\n\n')
                            else:
                                self.result_signal.emit('通道' + str(ec + 2) + '的通道灯存在问题。\n\n')
            # self.testByIndex(33, 'DI')
            # 等待子线程执行结束
            even_thread.join()
            self.testByIndex(33, 'DI', 0)

        DI_endTime = time.time()
        DI_testTime = round((DI_endTime - DI_startTime), 3)
        if self.isDIPassTest == False:
            self.pauseOption()
            if not self.is_running:
                return False
            self.messageBox_signal.emit(['DI通道指示灯结果', 'DI通道测试未通过！'])
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(self.HORIZONTAL_LINE + 'DI通道测试未通过！\n' + self.HORIZONTAL_LINE)
            self.item_signal.emit([4, 2, 2, f'{DI_testTime}'])
        elif self.isDIPassTest == True:
            # self.pauseOption()
            # if not self.is_running:
            #     return False
            # self.messageBox_signal.emit(['DI通道指示灯结果', 'DI通道测试全部通过！'])
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(self.HORIZONTAL_LINE + 'DI通道测试全部通过！\n' + self.HORIZONTAL_LINE)
            self.item_signal.emit([4, 2, 1, f'{DI_testTime}'])

        return True

    def testDO(self):
        DO_startTime = time.time()
        if self.isTestRunErr or self.isTestCANRunErr:
            # 进入指示灯测试模式
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("--------------进入 LED TEST模式-------------\n\n")
            # print("--------------进入 LED TEST模式-------------")
            # self.channelZero()
            self.m_transmitData[0] = 0x23
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x00
            self.m_transmitData[4] = 0x53
            self.m_transmitData[5] = 0x54
            self.m_transmitData[6] = 0x41
            self.m_transmitData[7] = 0x52
            isLEDTest, whatEver = CAN_option.transmitCAN((0x600 + int(self.CANAddr_DO)), self.m_transmitData, 1)
            if isLEDTest:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit("-----------进入 LED TEST 模式成功-----------\n")
                # print("-----------进入 LED TEST 模式成功-----------" + self.HORIZONTAL_LINE)

                if self.isTestRunErr:
                    if not self.testRunErr(int(self.CANAddr_DO), 'DO'):
                        self.pass_signal.emit(False)
                if self.isTestCANRunErr:
                    if not self.testCANRunErr(int(self.CANAddr_DO)):
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
                bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + self.CANAddr_DO), self.m_transmitData,
                                                                       1)
                if bool_transmit:
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    self.result_signal.emit("成功退出 LED TEST 模式。\n" + self.HORIZONTAL_LINE)
                    # print("成功退出 LED TEST 模式\n" + self.HORIZONTAL_LINE)
                else:
                    self.result_signal.emit("退出 LED TEST 模式失败！\n" + self.HORIZONTAL_LINE)
            else:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit("-----------进入 LED TEST 模式失败-----------\n\n")
                # print("-----------进入 LED TEST 模式失败-----------\n\n")
                self.item_signal.emit([0, 2, 2, '进入模式失败'])
                self.item_signal.emit([1, 2, 2, '进入模式失败'])

        self.item_signal.emit([4, 1, 0, ''])
        self.DODataCheck = [True for i in range(32)]  # 预留了最多32个通道的测试结果
        for i in range(self.loop_num):
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f"第{i + 1}次循环" + self.HORIZONTAL_LINE)
            CAN_option.Can_DLL.VCI_ClearBuffer(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX, CAN_option.CAN_INDEX)
            time.sleep(0.2)
            self.pauseOption()
            if not self.is_running:
                return False

            # 奇数通道检测
            self.pauseOption()
            if not self.is_running:
                return False
            # self.messageBox_oneButton_signal.emit(['操作提示', '请观察待测模块通道指示灯亮灭情况。'])
            # reply = self.result_queue.get()
            # if reply == QMessageBox.AcceptRole:
            import threading
            odd_thread = threading.Thread(target=self.oddChannelTest,args=('DO',))
            odd_thread.start()
            image_QNP0016_LED = self.current_dir + '/QNP0016_odd.gif'
            self.gif_messageBox_signal.emit(
                [f'{self.module_type}通道检测', f'{self.module_type}通道指示灯是否如图所示闪烁？', image_QNP0016_LED])
            reply = self.result_queue.get()

            self.stopChannel = reply
            if reply == QMessageBox.AcceptRole:
                self.isDOPassTest &= True
            else:
                self.isDOPassTest &= False

                self.checkBox_messageBox_signal.emit(['故障通道选择', '请选择存在故障的通道灯。', 'odd'])
                messageList = self.result_queue.get()
                reply = messageList[0]
                channelLED_odd = messageList[1]
                if reply == QMessageBox.AcceptRole:
                    for ec in range(1, 16, 2):
                        self.channelLED[ec] &= channelLED_odd[ec]
                        if not channelLED_odd[ec]:
                            if ec < 8:
                                self.result_signal.emit('通道0' + str(ec) + '的通道灯存在问题。\n\n')
                            else:
                                self.result_signal.emit('通道' + str(ec + 2) + '的通道灯存在问题。\n\n')
            # 等待子线程执行结束
            odd_thread.join()
            self.testByIndex(33, 'DO', 0)

            # 偶数通道检测
            self.stopChannel = 2
            self.pauseOption()
            if not self.is_running:
                return False
            # self.messageBox_oneButton_signal.emit(['操作提示', '请观察待测模块通道指示灯亮灭情况。'])
            # reply = self.result_queue.get()
            # if reply == QMessageBox.AcceptRole:
            even_thread = threading.Thread(target=self.evenChannelTest, args=('DO',))
            even_thread.start()
            image_QNP0016_LED = self.current_dir + '/QNP0016_even.gif'
            self.gif_messageBox_signal.emit(
                [f'{self.module_type}通道检测', f'{self.module_type}通道指示灯是否如图所示闪烁？',
                 image_QNP0016_LED])
            reply = self.result_queue.get()

            self.stopChannel = reply
            if reply == QMessageBox.AcceptRole:
                self.isDOPassTest &= True
            else:
                self.isDOPassTest &= False
                self.checkBox_messageBox_signal.emit(['故障通道选择', '请选择存在故障的通道灯。', 'even'])
                messageList = self.result_queue.get()
                reply = messageList[0]
                channelLED_even = messageList[1]
                if reply == QMessageBox.AcceptRole:
                    for ec in range(0, 15, 2):
                        self.channelLED[ec] &= channelLED_even[ec]
                        if not channelLED_even[ec]:
                            if ec < 7:
                                self.result_signal.emit('通道0' + str(ec) + '的通道灯存在问题。\n\n')
                            else:
                                self.result_signal.emit('通道' + str(ec + 2) + '的通道灯存在问题。\n\n')
            # self.testByIndex(33, 'DI')
            # 等待子线程执行结束
            even_thread.join()
            self.testByIndex(33, 'DO', 0)

        DO_endTime = time.time()
        DO_testTime = round((DO_endTime - DO_startTime), 3)
        if self.isDOPassTest == False:
            self.pauseOption()
            if not self.is_running:
                return False
            self.messageBox_signal.emit(['DO通道指示灯结果', f'DO通道测试未通过！'])
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(self.HORIZONTAL_LINE + 'DO通道测试未通过！\n' + self.HORIZONTAL_LINE)
            self.item_signal.emit([4, 2, 2, f'{DO_testTime}'])
        elif self.isDOPassTest == True:
            # self.pauseOption()
            # if not self.is_running:
            #     return False
            # self.messageBox_signal.emit(['DO通道指示灯结果', f'DO通道测试全部通过！'])
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(self.HORIZONTAL_LINE + 'DO通道测试全部通过！\n' + self.HORIZONTAL_LINE)
            self.item_signal.emit([4, 2, 1, f'{DO_testTime}'])

        return True

    def testByIndex(self, index, type: str,times:int=1):
        """
        :param index:
        :param type:
        :param times: 打印数据的标志，只有times==1时会打印
        :return:
        """
        self.sendTestDataToDO(index)
        self.hex_m_transmitData = ['','','','']
        if index == 32:
            self.pauseOption()
            if not self.is_running:
                return False
            if times ==1:
                self.result_signal.emit(
                    f'{1}.发送的数据：{hex(self.m_transmitData[0])}  {hex(self.m_transmitData[1])}  {hex(self.m_transmitData[2])}'
                    f'{hex(self.m_transmitData[3])}\n\n')
        elif index == 33:
            self.pauseOption()
            if not self.is_running:
                return False
            if times == 1:
                self.result_signal.emit(
                    f'{self.m_Channels + 2}.发送的数据：{hex(self.m_transmitData[0])}  '
                    f'{hex(self.m_transmitData[1])}  {hex(self.m_transmitData[2])}  '
                    f'{hex(self.m_transmitData[3])}\n\n')
        else:
            self.pauseOption()
            if not self.is_running:
                return False
            if times == 1:
                self.result_signal.emit(f'{index + 2}.发送的数据：{hex(self.m_transmitData[0])}  '
                                        f'{hex(self.m_transmitData[1])}  {hex(self.m_transmitData[2])}  '
                                        f'{hex(self.m_transmitData[3])}\n\n')

        now_time = time.time()
        while True:
            if (time.time() - now_time) * 1000 > 200:
                self.pauseOption()
                if not self.is_running:
                    return False
                if times == 1:
                    self.result_signal.emit(f'  接收的数据：{hex(self.m_receiveData[0])}  '
                                            f'{hex(self.m_receiveData[1])}  {hex(self.m_receiveData[2])}  '
                                            f'{hex(self.m_receiveData[3])}  收发不一致！\n\n')
                self.errorRece = [0,0,0,0]
                self.errorRece[0] = hex(self.m_receiveData[0])
                self.errorRece[1] = hex(self.m_receiveData[1])
                self.errorRece[2] = hex(self.m_receiveData[2])
                self.errorRece[3] = hex(self.m_receiveData[3])

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

                self.pauseOption()
                if not self.is_running:
                    return False
                return False
                # self.messageBox_signal.emit(['警告', '检测到收发不一致！'])
                # reply = self.result_queue.get()
                # if reply == QMessageBox.AcceptRole:
                #     return False
                # else:
                #     return False
            # 清理接收缓存区
            self.clearList(self.m_receiveData)
            bool_receive, self.m_can_obj = CAN_option.receiveCANbyID((0x180 + self.CANAddr_DI), self.TIME_OUT, 1)
            self.m_receiveData = self.m_can_obj.Data

            if bool_receive == False:
                self.pauseOption()
                if not self.is_running:
                    return False
                if times == 1:
                    self.result_signal.emit('\n  接收数据超时！\n\n')
                # print('接收数据超时！')
                if type == 'DO':
                    self.isDOPassTest = False
                    self.DODataCheck[index] = False
                if type == 'DI':
                    self.isDIPassTest = False
                    self.DIDataCheck[index] = False
                break
            elif self.m_transmitData[0] == self.m_receiveData[0] and \
                    self.m_transmitData[1] == self.m_receiveData[1] and \
                    self.m_transmitData[2] == self.m_receiveData[2] and \
                    self.m_transmitData[3] == self.m_receiveData[3]:
                self.pauseOption()
                if not self.is_running:
                    return False
                if times == 1:
                    self.result_signal.emit(f'  接收的数据：{hex(self.m_receiveData[0])}  {hex(self.m_receiveData[1])}  '
                                            f'{hex(self.m_receiveData[2])}  {hex(self.m_receiveData[3])}  收发一致\n\n')
                break
        return True


    def oddChannelTest(self,type:str):
        # 奇数通道灯闪烁5s
        startTime = time.time()
        times = 1
        while True:
            if self.stopChannel == QMessageBox.AcceptRole or self.stopChannel == QMessageBox.RejectRole:
                self.stopChannel= 2
                self.testByIndex(34, type ,times)
                break
            if self.testByIndex(34, type ,times):
                time.sleep(0.3)
            else:
                if times == 1:
                    bin_0 = bin(int('0xaa', 16)&int(self.errorRece[0], 16))[2:]
                    if len(bin_0)<8:
                        for b_num in range(8-len(bin_0)):
                            bin_0='0'+bin_0
                    bin_0=bin_0[::-1]
                    self.result_signal.emit(f'通道0-7：{bin_0}\n')
                    if self.m_Channels == 16:
                        bin_1 = bin(int('0xaa', 16)&int(self.errorRece[1], 16))[2:]
                        if len(bin_1) < 8:
                            for b_num in range(8 - len(bin_1)):
                                bin_1 = '0' + bin_1
                        bin_1 = bin_1[::-1]
                        self.result_signal.emit(f'通道10-17：{bin_1}\n')
                        bin_all = bin_0+bin_1
                        for odd in range(1, self.m_Channels + 1, 2):
                            if bin_all[odd] == '0':
                                if odd > 7:
                                    odd_real = odd+2
                                else:
                                    odd_real = odd
                                self.result_signal.emit(f'通道{odd_real}接收数据为{bin_all[odd]}，存在问题。\n\n')
                                if type == 'DI':
                                    self.isDIPassTest &= False
                                    self.DIDataCheck[odd] &= False
                                elif type == 'DO':
                                    self.isDOPassTest &= False
                                    self.DODataCheck[odd] &= False
                            else:
                                if type == 'DI':
                                    self.isDIPassTest &= True
                                    self.DIDataCheck[odd] &= True
                                elif type == 'DO':
                                    self.isDOPassTest &= True
                                    self.DODataCheck[odd] &= True

            if self.testByIndex(33, type ,times):
                time.sleep(0.3)
            times += 1
                # break
                # if self.m_Channels == 31:
                #     self.result_signal.emit(f'通道10-17：{bin(self.m_transmitData[1] & self.m_receiveData[1])[2:]}\n')
                #     self.result_signal.emit(f'通道20-27：{bin(self.m_transmitData[2] & self.m_receiveData[2])[2:]}\n')
                #     self.result_signal.emit(f'通道30-37：{bin(self.m_transmitData[3] & self.m_receiveData[3])[2:]}\n')


    def evenChannelTest(self,type:str):
        # 偶数通道灯闪烁5s
        startTime = time.time()
        times = 1
        while True:
            # if time.time() - startTime > 5:
            if self.stopChannel == QMessageBox.AcceptRole or self.stopChannel == QMessageBox.RejectRole:
                self.stopChannel= 2
                self.testByIndex(35, type ,times)
                break
            if self.testByIndex(35, type ,times):
                time.sleep(0.3)
            else:
                if times == 1:
                    bin_0 = bin(int('0x55', 16)&int(self.errorRece[0], 16))[2:]
                    if len(bin_0)<8:
                        for b_num in range(8-len(bin_0)):
                            bin_0='0'+bin_0
                    bin_0=bin_0[::-1]
                    self.result_signal.emit(f'通道0-7：{bin_0}\n')
                    if self.m_Channels == 16:
                        bin_1 = bin(int('0x55', 16)&int(self.errorRece[1], 16))[2:]
                        if len(bin_1) < 8:
                            for b_num in range(8 - len(bin_1)):
                                bin_1 = '0' + bin_1
                        bin_1 = bin_1[::-1]
                        self.result_signal.emit(f'通道10-17：{bin_1}\n')
                        bin_all = bin_0+bin_1
                        for even in range(0, self.m_Channels, 2):
                            if bin_all[even] == '0':
                                if even > 6:
                                    even_real = even+2
                                else:
                                    even_real = even
                                self.result_signal.emit(f'通道{even_real}接收数据为{bin_all[even]}，存在问题。\n\n')
                                if type == 'DI':
                                    self.isDIPassTest &= False
                                    self.DIDataCheck[even] &= False
                                elif type == 'DO':
                                    self.isDOPassTest &= False
                                    self.DODataCheck[even] &= False
                            else:
                                if type == 'DI':
                                    self.isDIPassTest &= True
                                    self.DIDataCheck[even] &= True
                                elif type == 'DO':
                                    self.isDOPassTest &= True
                                    self.DODataCheck[even] &= True

            if self.testByIndex(33, type ,times):
                time.sleep(0.3)
            times += 1

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
        CAN_option.transmitCAN(0x200 + self.CANAddr_DO, self.m_transmitData, 1)

    def testRunErr(self, addr, type):
        self.testNum -= 1

        self.isLEDRunOK = True

        self.isLEDErrOK = True
        self.isLEDPass = True
        self.pauseOption()
        if not self.is_running:
            return False
        self.item_signal.emit([0, 1, 0, ''])

        runStart_time = time.time()

        self.pauseOption()
        if not self.is_running:
            return False
        self.result_signal.emit("1.进行 LED RUN 测试\n\n")
        # print("1.进行 LED RUN 测试\n\n")
        self.clearList(self.m_transmitData)
        self.m_transmitData[0] = 0x2f
        self.m_transmitData[1] = 0xf6
        self.m_transmitData[2] = 0x5f
        self.m_transmitData[3] = 0x01
        self.m_transmitData[4] = 0x01
        bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData, 1)
        runEnd_time = time.time()
        runTest_time = round(runEnd_time - runStart_time, 2)
        time.sleep(0.1)

        if type == 'DI':
            image_RUN = self.current_dir + '/ET1600_RUN.png'
        elif type == 'DO':
            image_RUN = self.current_dir + '/DO_RUN.png'
        self.pic_messageBox_signal.emit(['检测RUN &ERROR', 'RUN指示灯是否如红框中所示点亮？', image_RUN])
        # self.messageBox_signal.emit(['检测RUN &ERROR', 'RUN指示灯是否点亮？'])
        reply = self.result_queue.get()
        if reply == QMessageBox.AcceptRole:
            self.runLED = True

            self.pauseOption()
            if not self.is_running:
                return False
            self.item_signal.emit([0, 2, 1, f'{runTest_time}'])
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("LED RUN 测试通过\n关闭LED RUN\n" + self.HORIZONTAL_LINE)
            # print("LED RUN 测试通过\n关闭LED RUN\n" + self.HORIZONTAL_LINE)
            self.clearList(self.m_transmitData)
            self.m_transmitData[0] = 0x2f
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x01
            self.m_transmitData[4] = 0x00
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData, 1)
            time.sleep(0.2)
            self.isLEDRunOK = True
            # self.result_signal.emit(f'self.isLEDRunOK:{self.isLEDRunOK}')
        elif reply == QMessageBox.RejectRole:
            self.runLED = False
            self.pauseOption()
            if not self.is_running:
                return False
            self.item_signal.emit([0, 2, 2, f'{runTest_time}'])

            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("LED RUN 测试不通过\n" + self.HORIZONTAL_LINE)
            # print("LED RUN 测试不通过\n" + self.HORIZONTAL_LINE)
            self.isLEDRunOK = False
            # self.result_signal.emit(f'self.isLEDRunOK:{self.isLEDRunOK}')
        self.isLEDPass = self.isLEDPass & self.isLEDRunOK

        errorStart_time = time.time()
        self.pauseOption()
        if not self.is_running:
            return False
        self.result_signal.emit("2.进行 LED ERROR 测试\n\n")
        # print("2.进行 LED ERROR 测试\n\n")
        self.pauseOption()
        if not self.is_running:
            return False
        self.item_signal.emit([1, 1, 0, ''])


        self.clearList(self.m_transmitData)
        self.m_transmitData[0] = 0x2f
        self.m_transmitData[1] = 0xf6
        self.m_transmitData[2] = 0x5f
        self.m_transmitData[3] = 0x02
        self.m_transmitData[4] = 0x01

        bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData, 1)
        errorEnd_time = time.time()
        errorTest_time = round(errorEnd_time - errorStart_time, 2)
        time.sleep(0.1)
        if type == 'DI':
            image_ERROR = self.current_dir + '/ET1600_ERROR.png'
        elif type == 'DO':
            image_ERROR = self.current_dir + '/DO_ERROR.png'
        self.pic_messageBox_signal.emit(['检测RUN &ERROR', 'ERROR指示灯是否如红框中所示点亮？', image_ERROR])
        reply = self.result_queue.get()

        if reply == QMessageBox.AcceptRole:
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
            self.item_signal.emit([1, 2, 1, f'{errorTest_time}'])
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("LED ERROR 测试通过\n关闭LED ERROR\n" + self.HORIZONTAL_LINE)
            # print("LED ERROR 测试通过\n关闭LED ERROR\n" + self.HORIZONTAL_LINE)
            self.clearList(self.m_transmitData)
            self.m_transmitData[0] = 0x2f
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x02
            self.m_transmitData[4] = 0x00
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData, 1)
            time.sleep(0.2)
            self.isLEDErrOK = True
        elif reply == QMessageBox.RejectRole:
            self.errorLED = False
            self.pauseOption()
            if not self.is_running:
                return False
            self.item_signal.emit([1, 2, 2, f'{errorTest_time}'])
            # self.itemOself.pauseOption()peration(mTable, 2, 2, 2, errorTest_time)
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("LED ERROR 测试不通过\n" + self.HORIZONTAL_LINE)
            # print("LED ERROR 测试不通过\n" + self.HORIZONTAL_LINE)
            self.isLEDErrOK = False
        self.isLEDPass = self.isLEDPass & self.isLEDErrOK

        return True

    def testCANRunErr(self, addr):
        self.testNum -= 1

        CANRunStart_time = time.time()
        self.pauseOption()
        if not self.is_running:
            return False
        self.item_signal.emit([2, 1, 0, ''])
        # 进入指示灯测试模式
        self.pauseOption()
        if not self.is_running:
            return False
        # self.result_signal.emit("--------------进入 LED TEST模式-------------\n\n")
        # #print("--------------进入 LED TEST模式-------------")
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
        # print(f'{self.module_2}地址:{0x600 + addr}')
        isLEDTest, whatEver = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData, 1)
        if isLEDTest:
            self.pauseOption()
            if not self.is_running:
                return False
            # self.result_signal.emit("-----------进入 LED TEST 模式成功-----------\n")
            # #print("-----------进入 LED TEST 模式成功-----------" + self.HORIZONTAL_LINE)
            # self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("1.进行 LED CAN_RUN 测试\n\n")
            # print("1.进行 LED CAN_RUN 测试\n\n")
            self.clearList(self.m_transmitData)
            self.m_transmitData[0] = 0x2f
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x03
            self.m_transmitData[4] = 0x01
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData, 1)
            CANRunEnd_time = time.time()
            CANRunTest_time = round(CANRunEnd_time - CANRunStart_time, 2)
            time.sleep(0.1)
            self.messageBox_signal.emit(['检测RUN &ERROR', 'CAN_RUN指示灯是否点亮（绿色）？'])
            reply = self.result_queue.get()
            # reply = QMessageBox.question(None, '检测CAN_RUN &CAN_ERROR', '指示灯是否点亮？',
            #                              QMessageBox.AcceptRole | QMessageBox.RejectRole,
            #                              QMessageBox.AcceptRole)
            if reply == QMessageBox.AcceptRole:
                self.CAN_runLED = True
                self.pauseOption()
                if not self.is_running:
                    return False
                self.item_signal.emit([2, 2, 1, CANRunTest_time])
                # self.itemOperation(mTable,3,2,1,CANRunTest_time)
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit("LED CAN_RUN 测试通过\n关闭LED CAN_RUN\n" + self.HORIZONTAL_LINE)
                # print("LED CAN_RUN 测试通过\n关闭LED CAN_RUN\n" + self.HORIZONTAL_LINE)
                self.clearList(self.m_transmitData)
                self.m_transmitData[0] = 0x2f
                self.m_transmitData[1] = 0xf6
                self.m_transmitData[2] = 0x5f
                self.m_transmitData[3] = 0x03
                self.m_transmitData[4] = 0x00
                bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData, 1)
                time.sleep(0.2)
                isLEDCANRunOK = True
            elif reply == QMessageBox.RejectRole:
                self.CAN_runLED = False
                self.pauseOption()
                if not self.is_running:
                    return False
                self.item_signal.emit([2, 2, 2, CANRunTest_time])
                # self.itemOperation(mTable, 3, 2, 2, CANRunTest_time)
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit("LED CAN_RUN 测试不通过\n" + self.HORIZONTAL_LINE)
                # print("LED CAN_RUN 测试不通过\n" + self.HORIZONTAL_LINE)
                isLEDCANRunOK = False

            CANErrStart_time = time.time()
            self.pauseOption()
            if not self.is_running:
                return False
            self.item_signal.emit([3, 1, 0, ''])
            # self.itemOperation(mTable, 4, 1, 0, '')
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("2.进行 LED CAN_ERROR 测试\n\n")
            # print("2.进行 LED CAN_ERROR 测试\n\n")
            self.clearList(self.m_transmitData)
            self.m_transmitData[0] = 0x2f
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x04
            self.m_transmitData[4] = 0x01
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData, 1)
            CANErrEnd_time = time.time()
            CANErrTest_time = round(CANErrEnd_time - CANErrStart_time, 2)
            time.sleep(0.1)
            self.messageBox_signal.emit(['检测RUN &ERROR', 'CAN_ERROR指示灯是否点亮(红色)？'])
            reply = self.result_queue.get()

            if reply == QMessageBox.AcceptRole:
                self.CAN_errorLED = True
                self.pauseOption()
                if not self.is_running:
                    return False
                self.item_signal.emit([3, 2, 1, CANErrTest_time])
                # self.itemOperation(mTable, 4, 2, 1, CANErrTest_time)
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit("LED CAN_ERROR 测试通过\n关闭LED CAN_ERROR\n" + self.HORIZONTAL_LINE)

                self.clearList(self.m_transmitData)
                self.m_transmitData[0] = 0x2f
                self.m_transmitData[1] = 0xf6
                self.m_transmitData[2] = 0x5f
                self.m_transmitData[3] = 0x04
                self.m_transmitData[4] = 0x00
                bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData, 1)
                time.sleep(0.2)
                isLEDCANErrOK = True
            elif reply == QMessageBox.RejectRole:
                self.CAN_errorLED = False
                self.pauseOption()
                if not self.is_running:
                    return False
                self.item_signal.emit([3, 2, 2, CANErrTest_time])

                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit("LED CAN_ERROR 测试不通过\n" + self.HORIZONTAL_LINE)

                isLEDCANErrOK = False
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("退出 LED TEST 模式\n" + self.HORIZONTAL_LINE)
            # print("退出 LED TEST 模式\n" + self.HORIZONTAL_LINE)
            self.clearList(self.m_transmitData)
            self.m_transmitData[0] = 0x23
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x00
            self.m_transmitData[4] = 0x45
            self.m_transmitData[5] = 0x58
            self.m_transmitData[6] = 0x49
            self.m_transmitData[7] = 0x54
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData, 1)

        else:
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("-----------进入 LED TEST 模式失败-----------\n\n")
            # print("-----------进入 LED TEST 模式失败-----------\n\n")
            self.item_signal.emit([2, 2, 2, '进入模式失败'])
            self.item_signal.emit([3, 2, 2, '进入模式失败'])

        return True

    def normal_writeValuetoAO(self, value):
        bool_all = True
        for i in range(self.m_Channels):
            self.m_transmitData = [0x2b, 0x11, 0x64, i + 1, (value & 0xff), ((value >> 8) & 0xff), 0x00, 0x00]
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + self.CANAddr_AO), self.m_transmitData, 1)
            bool_all = bool_all & bool_transmit

        return bool_all

    def channelZero(self):
        self.pauseOption()
        if not self.is_running:
            return False
        self.result_signal.emit(f'对所有通道值进行归零处理' + self.HORIZONTAL_LINE)
        # #print('1111111111111111')
        isZero = self.normal_writeValuetoAO(0)
        # #print('2222222222222222')
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

        # 填写通道状态

        for i in range(self.m_Channels):
            if i < 16:
                if (self.DIDataCheck[i] and self.channelLED[i]):
                    sheet.write(self.DI_row + 1, 3 + i, '√', pass_style)
                elif not (self.DIDataCheck[i] and self.channelLED[i]):
                    sheet.write(self.DI_row + 1, 3 + i, '×', fail_style)
                    if i < 8:
                        if not self.DIDataCheck[i]:
                            self.errorNum += 1
                            self.errorInf += f'\n{self.errorNum})通道0{i}故障 '
                        if not self.channelLED[i]:
                            self.errorNum += 1
                            self.errorInf += f'\n{self.errorNum})通道0{i}灯故障 '
                    else:
                        if not self.DIDataCheck[i]:
                            self.errorNum += 1
                            self.errorInf += f'\n{self.errorNum})通道{i + 2}故障 '
                        if not self.channelLED[i]:
                            self.errorNum += 1
                            self.errorInf += f'\n{self.errorNum})通道{i + 2}灯故障 '
            else:
                if self.DIDataCheck[i]:
                    sheet.write(self.DI_row + 3, i - 13, '√', pass_style)
                elif not self.DIDataCheck[i]:
                    sheet.write(self.DI_row + 3, i - 13, '×', fail_style)
                    self.errorNum += 1
                    self.errorInf += f'\n{self.errorNum})通道{i + 1} 灯未亮 '

        self.isDIPassTest = (((((self.isDIPassTest & self.isLEDRunOK) & self.isLEDErrOK) &
                               self.CAN_runLED) & self.CAN_errorLED) & self.appearance)

        all_row = 2 + 9 + 4 + 4 + 2 + 2  # 常规+CPU + DI + DO + AI + AO
        if self.isDIPassTest and self.testNum == 0:
            name_save = '合格'
            sheet.write(self.generalTest_row + all_row, 4, '■ 合格', pass_style)
            self.label_signal.emit(['pass', '全部通过'])
            self.print_signal.emit(
                [f'/{name_save}{self.module_type}_{self.module_sn}', 'PASS', ''])

        if self.isDIPassTest and self.testNum > 0:
            name_save = '部分合格'
            sheet.write(self.generalTest_row + all_row, 4, '■ 部分合格', pass_style)
            sheet.write(self.generalTest_row + all_row, 6,
                        '------------------ 注意：有部分项目未测试！！！ ------------------', warning_style)
            self.label_signal.emit(['testing', '部分通过'])
            self.print_signal.emit(
                [f'/{name_save}{self.module_type}_{self.module_sn}', 'PASS', ''])
            # self.label.setStyleSheet(self.testState_qss['testing'])
            # self.label.setText('部分通过')
        elif not self.isDIPassTest:
            name_save = '不合格'
            sheet.write(self.generalTest_row + all_row + 1, 4, '■ 不合格', fail_style)
            sheet.write(self.generalTest_row + all_row + 1, 6, f'不合格原因：{self.errorInf}', warning_style)
            self.label_signal.emit(['fail', '未通过'])
            self.print_signal.emit(
                [f'/{name_save}{self.module_type}_{self.module_sn}', 'FAIL', self.errorInf])
            # self.label.setStyleSheet(self.testState_qss['fail'])
            # self.label.setText('未通过')
        # book.save(self.saveDir + f'/{name_save}{self.module_type}_{self.module_sn}.xls')
        self.saveExcel_signal.emit([book, f'/{name_save}{self.module_type}_{self.module_sn}.xls'])

    # @abstractmethod
    def fillInDOData(self, station, book, sheet):
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
                if (self.DODataCheck[i] and self.channelLED[i]):
                    sheet.write(self.DO_row + 1, 3 + i, '√', pass_style)
                elif not (self.DODataCheck[i] and self.channelLED[i]):
                    sheet.write(self.DO_row + 1, 3 + i, '×', fail_style)
                    if i < 8:
                        if not self.DODataCheck[i]:
                            self.errorNum += 1
                            self.errorInf += f'\n{self.errorNum})通道0{i}故障 '
                        if not self.channelLED[i]:
                            self.errorNum += 1
                            self.errorInf += f'\n{self.errorNum})通道0{i}灯故障 '
                    else:
                        if not self.DODataCheck[i]:
                            self.errorNum += 1
                            self.errorInf += f'\n{self.errorNum})通道{i + 2}故障 '
                        if not self.channelLED[i]:
                            self.errorNum += 1
                            self.errorInf += f'\n{self.errorNum})通道{i + 2}灯故障 '
            else:
                if self.DODataCheck[i]:
                    sheet.write(self.DO_row + 3, i - 13, '√', pass_style)
                elif not self.DODataCheck[i]:
                    sheet.write(self.DO_row + 3, i - 13, '×', fail_style)
                    self.errorNum += 1
                    self.errorInf += f'\n{self.errorNum}.通道{i + 1}灯未亮 '

        self.isDOPassTest = (((((self.isDOPassTest & self.isLEDRunOK) & self.isLEDErrOK) &
                               self.CAN_runLED) & self.CAN_errorLED) & self.appearance)

        all_row =2 + 9 + 4 + 4 + 2 + 2  # 常规+CPU + DI + DO + AI + AO
        if self.isDOPassTest and self.testNum == 0:
            name_save = '合格'
            sheet.write(self.generalTest_row + all_row, 4, '■ 合格', pass_style)
            self.label_signal.emit(['pass', '全部通过'])
            self.print_signal.emit([f'/{name_save}{self.module_type}_{self.module_sn}', 'PASS', ''])

        if self.isDOPassTest and self.testNum > 0:
            name_save = '部分合格'
            sheet.write(self.generalTest_row + all_row, 4, '■ 部分合格', pass_style)
            sheet.write(self.generalTest_row + all_row, 6,
                        '------------------ 注意：有部分项目未测试！！！ ------------------', warning_style)
            self.label_signal.emit(['testing', '部分通过'])
            self.print_signal.emit([f'/{name_save}{self.module_type}_{self.module_sn}', 'PASS', ''])

        elif not self.isDOPassTest:
            name_save = '不合格'
            sheet.write(self.generalTest_row + all_row + 1, 4, '■ 不合格', fail_style)
            sheet.write(self.generalTest_row + all_row + 1, 6, f'不合格原因：{self.errorInf}', fail_style)
            self.label_signal.emit(['fail', '未通过'])
            self.print_signal.emit(
                [f'/{name_save}{self.module_type}_{self.module_sn}', 'FAIL', self.errorInf])

        self.saveExcel_signal.emit([book, f'/{name_save}{self.module_type}_{self.module_sn}.xls'])

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

