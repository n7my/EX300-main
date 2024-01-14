# -*- coding: utf-8 -*-

import time

from PyQt5.QtCore import QThread, pyqtSignal
from main_logic import *
import threading
from  CAN_option import *
import struct
import otherOption

class AIThread(QObject):
    #result_signal：在运行日志框中显示信息
    result_signal = pyqtSignal(str)
    #item_signal：修改表格状态信号
    item_signal = pyqtSignal(list)
    pass_signal = pyqtSignal(bool)
    # RunErr_signal = pyqtSignal(int)
    # CANRunErr_signal = pyqtSignal(int)
    #messageBox_signal：显示对话框信号
    messageBox_signal= pyqtSignal(list)
    pic_messageBox_signal = pyqtSignal(list)
    #allFinished_signal：所有测试结束信号
    allFinished_signal = pyqtSignal()
    label_signal = pyqtSignal(list)
    saveExcel_signal = pyqtSignal(list)
    print_signal = pyqtSignal(list)

    HORIZONTAL_LINE = "\n------------------------------------------------------------------------------------------------------------\n\n"
    #                -/+10V       0-10V       0-5V        1-5V       4-20mA      0-20mA      -/+5V
    AORangeArray = [0x0b, 0x00, 0xe8, 0x03, 0xe9, 0x03, 0xea, 0x03, 0x15, 0x00, 0xeb, 0x03, 0xec, 0x03]
    AIRangeArray = [0x29, 0x00, 0x2a, 0x00, 0x10, 0x27, 0x11, 0x27, 0x33, 0x00, 0x34, 0x00, 0x12, 0x27]
    # AO量程
    #                    -/+10V        -/+5V        0-5V       0-10V       1-5V
    vol_AORangeArray = [0x0b, 0x00, 0xec, 0x03, 0xe9, 0x03, 0xe8, 0x03, 0xea, 0x03]
    #                     4-20mA      0-20mA
    cur_AORangeArray = [0x15, 0x00, 0xeb, 0x03]
    # AI量程
    #                    -/+10V        -/+5V        0-5V       0-10V       1-5V
    vol_AIRangeArray = [0x29, 0x00, 0x12, 0x27, 0x10, 0x27, 0x2a, 0x00, 0x11, 0x27]
    #                     4-20mA      0-20mA
    cur_AIRangeArray = [0x33, 0x00, 0x34, 0x00]

    vol_name_array = ['电压+/-10V', '电压+/-5V', '电压0-5V', '电压0-10V', '电压1-5V']
    cur_name_array = ['电流4-20mA', '电流0-20mA']

    # 4mA-20mA的理论值
    currentTheory_0420 = [0x6c00, (int)(0x6c00 * 0.75), (int)(0x6c00 * 0.5), (int)(0x6c00 * 0.25), 0]  # 0x6c00 =>27648
    arrayCur_0420 = ["20mA测试", "16mA测试", "12mA测试", "8mA测试", "4mA测试"]
    # 4mA-20mA的rawData
    highCurrent_0420 = 54162
    lowCurrent_0420 = 10832

    # 0mA-20mA的理论值
    currentTheory_0020 = [0x6c00, (int)(0x6c00 * 0.75), (int)(0x6c00 * 0.5), (int)(0x6c00 * 0.25), 0]  # 0x6c00 =>27648
    arrayCur_0020 = ["20mA测试", "15mA测试", "10mA测试", "5mA测试", "0mA测试"]
    # 0mA-20mA的rawData
    highCurrent_0020 = 59849
    lowCurrent_0020 = 32768

    # -10v~10v的理论值
    voltageTheory_1010 = [0x6c00, 0x3600, 0x00, -13824, -27648]  # 0x6c00 =>27648
    arrayVol_1010 = ["10V测试", "5V测试", "0V测试", "-5V测试", "-10V测试"]
    # -10v~10v的rawData
    highVoltage_1010 = 59849
    lowVoltage_1010 = 5687

    # -5v~5v的理论值
    voltageTheory_0505 = [0x6c00, 0x3600, 0x00, -13824, -27648]  # 0x6c00 =>27648
    arrayVol_0505 = ["5V测试", "2.5V测试", "0V测试", "-2.5V测试", "-5V测试"]
    # -5v~5v的rawData
    highVoltage_0505 = 59849
    lowVoltage_0505 = 5687

    # 0v-5v的理论值
    voltageTheory_0005 = [0x6c00, (int)(0x6c00 * 0.75), (int)(0x6c00 * 0.5), (int)(0x6c00 * 0.25), 0]  # 0x6c00 =>27648
    arrayVol_0005 = ["5V测试", "3.75V测试", "2.50V测试", "1.25V测试", "0V测试"]
    # 0v-5v的rawData
    highVoltage_0005 = 59849
    lowVoltage_0005 = 32768

    # 0v-10v的理论值
    voltageTheory_0010 = [0x6c00, (int)(0x6c00 * 0.75), (int)(0x6c00 * 0.5), (int)(0x6c00 * 0.25), 0]  # 0x6c00 =>27648
    arrayVol_0010 = ["10V测试", "7.5V测试", "5V测试", "2.5V测试", "0V测试"]
    # 0v-10v的rawData
    highVoltage_0010 = 59849
    lowVoltage_0010 = 32768

    # 1v-5v的理论值
    voltageTheory_0105 = [0x6c00, (int)(0x6c00 * 0.75), (int)(0x6c00 * 0.5), (int)(0x6c00 * 0.25), 0]  # 0x6c00 =>27648
    arrayVol_0105 = ["5V测试", "4V测试", "3V测试", "2V测试", "1V测试"]
    # 1v-5v的rawData
    highVoltage_0105 = 54162
    lowVoltage_0105 = 10832

    # voltageTheory_1010 = [0x6c00, 0x3600, 0x00, 0xca00, -27648]  # 0x6c00 =>27648
    voltageTheory_1010 = [0x6c00, 0x3600, 0x00, -13824, -27648]  # 0x6c00 =>27648
    arrayVol_1010 = ["10V测试", "5V测试", "0V测试", "-5V测试", "-10V测试"]

    vol_high_standardValue_array = [367002, 314573, 314573, 367002, 314573]
    cur_high_standardValue_array = [314777, 314777]
    vol_low_standardValue_array = [157286, 209715, 262144, 262144, 272630]
    cur_low_standardValue_array = [272671, 262144]
    vol_error_value = [841, 841, 212, 422, 170]
    cur_high_error_value = [636, 678]
    cur_low_error_value = [265, 214]
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
    # 5个量程 -> 每个量程5个点 -> 每个点4个通道
    volReceValue = [
        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-']],
        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-']],
        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-']],
        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-']],
        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-']]]

    volPrecision = [
        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-']],
        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-']],
        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-']],
        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-']],
        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-']]]
    # 2个量程 -> 每个量程5个点 -> 每个点4个通道
    curReceValue = [
        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-']],
        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-']]]

    curPrecision = [
        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-']],
        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-']]]

    # AI测试是否通过
    isAIPassTest = True
    isAIVolPass = True
    isAICurPass = True

    isLEDRunOK = True
    isLEDErrOK = True
    isLEDPass = True

    isLEDCANRunOK = True
    isLEDCANErrOK = True
    isLEDCANPass = True

    # 发送的数据
    ubyte_array_transmit = c_ubyte * 8
    m_transmitData = ubyte_array_transmit(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
    #主副线程弹窗结果
    result_queue = 0

    CAN_errorLED = True
    CAN_runLED = True
    errorLED = True
    runLED = True

    #是否生成检测报告
    isExcel = True
    def __init__(self,inf_AIlist:list,result_queue,appearance):
        super().__init__()
        self.result_queue = result_queue
        self.is_running = True
        self.is_pause = False
        self.appearance = appearance
        self.mTable = inf_AIlist[0][0]
        self.module_1 = inf_AIlist[0][1]
        self.module_2 = inf_AIlist[0][2]
        self.testNum = inf_AIlist[0][3]

        # 获取产品信息
        self.module_type = inf_AIlist[1][0]
        self.module_pn = inf_AIlist[1][1]
        self.module_sn = inf_AIlist[1][2]
        self.module_rev = inf_AIlist[1][3]
        self.m_Channels = int(inf_AIlist[1][4])
        self.AI_Channels = self.m_Channels

        # 获取CAN地址
        self.CANAddr_AI = inf_AIlist[2][0]
        self.CANAddr_AO = inf_AIlist[2][1]
        self.CANAddr_relay = inf_AIlist[2][2]

        # 获取附加参数
        self.baud_rate = inf_AIlist[3][0]
        self.waiting_time = inf_AIlist[3][1]
        self.receive_num = inf_AIlist[3][2]
        self.saveDir = inf_AIlist[3][3]

        # 获取标定信息
        self.isCalibrate = inf_AIlist[4][0]
        self.isCalibrateVol = inf_AIlist[4][1]
        self.isCalibrateCur = inf_AIlist[4][2]
        self.oneOrAll = inf_AIlist[4][3]

        # 获取检测信息
        self.isTest = inf_AIlist[5][0]
        self.isAITestVol = inf_AIlist[5][1]
        self.isAITestCur = inf_AIlist[5][2]
        self.isTestCANRunErr = inf_AIlist[5][3]
        self.isTestRunErr = inf_AIlist[5][4]
        # print(f'inf_AIlist[5] = {inf_AIlist[5]}')
        self.errorNum = 0
        self.errorInf = ''
        self.pause_num = 1
        errorNum = 0

        self.current_dir = os.getcwd().replace('\\', '/') + "/_internal"


    def AIOption(self):
        self.isExcel = True
        # 总线初始化
        try:
            time_online = time.time()
            while True:
                QApplication.processEvents()
                if (time.time() - time_online) * 1000 > 2000:
                    self.pauseOption()
                    if not self.is_running:
                        # 后续测试全部取消
                        self.isTest = False
                        self.isCalibrate = False
                        self.isExcel = False
                        break
                    self.result_signal.emit(f'错误：总线初始化超时！' + self.HORIZONTAL_LINE)
                    QMessageBox.critical(None, '错误', '总线初始化超时！请检查CAN分析仪或各设备是否正确连接',
                                         QMessageBox.AcceptRole |
                                         QMessageBox.RejectRole, QMessageBox.AcceptRole)
                    # 后续测试全部取消
                    self.isTest = False
                    self.isCalibrate = False
                    self.isExcel = False
                    break
                CANAddr_array = [self.CANAddr_AO, self.CANAddr_AI, self.CANAddr_relay, self.CANAddr_relay + 1]
                module_array = [self.module_1, self.module_2, 'QR0016#1', 'QR0016#2']
                bool_online, eID = otherOption.isModulesOnline(CANAddr_array,module_array,self.waiting_time)
                if bool_online:
                    self.pauseOption()
                    if not self.is_running:
                        # 后续测试全部取消
                        self.isTest = False
                        self.isCalibrate = False
                        self.isExcel = False
                        break
                    self.result_signal.emit(f'总线初始化成功！' + self.HORIZONTAL_LINE)
                    # 点亮右扩IO设备的RUN灯
                    self.clearList(self.m_transmitData)
                    self.m_transmitData[0] = 0x52
                    self.m_transmitData[2] = 0x01
                    for led_run in range(2, 4):
                        self.m_transmitData[1] = led_run
                        bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x000, self.m_transmitData, 1)
                    break
                else:
                    self.pauseOption()
                    if not self.is_running:
                        # 后续测试全部取消
                        self.isTest = False
                        self.isCalibrate = False
                        self.isExcel = False
                        break
                    if eID != 0:
                        self.result_signal.emit(f'错误：未发现{module_array[eID - 1]}' + self.HORIZONTAL_LINE)
                    # if eID == 0:
                    #     self.result_signal.emit(f'错误：未发现{self.module_1}' + self.HORIZONTAL_LINE)
                    # elif eID == 1:
                    #     self.result_signal.emit(f'错误：未发现{self.module_2}' + self.HORIZONTAL_LINE)
                    # elif eID == 3:
                    #     self.result_signal.emit(f'错误：未发现继电器1' + self.HORIZONTAL_LINE)
                    # elif eID == 7:
                    #     self.result_signal.emit(f'错误：未发现继电器2' + self.HORIZONTAL_LINE)
                    self.result_signal.emit(f'错误：总线初始化失败！再次尝试初始化。' + self.HORIZONTAL_LINE)

            self.result_signal.emit('模块在线检测结束！' + self.HORIZONTAL_LINE)

        except:
            self.pauseOption()
            if not self.is_running:
                # 后续测试全部取消
                self.isTest = False
                self.isCalibrate = False
                self.isExcel = False

            QMessageBox(QMessageBox.Critical, '错误提示', '总线初始化异常，请检查设备').exec_()
            # 捕获异常并输出详细的错误信息
            traceback.print_exc()
            # 后续测试全部取消
            self.isTest = False
            self.isCalibrate = False
            self.isExcel = False
            self.result_signal.emit('模块在线检测结束！' + self.HORIZONTAL_LINE)
        #开始测试
        if self.isTest:
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
                isLEDTest, whatEver = CAN_option.transmitCAN((0x600 + int(self.CANAddr_AI)), self.m_transmitData,1)
                if isLEDTest:
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    self.result_signal.emit("-----------进入 LED TEST 模式成功-----------\n")
                    # print("-----------进入 LED TEST 模式成功-----------" + self.HORIZONTAL_LINE)

                    if self.isTestRunErr:
                        if not self.testRunErr(int(self.CANAddr_AI)):
                            self.pass_signal.emit(False)

                    if self.isTestCANRunErr:
                        if not self.testCANRunErr(int(self.CANAddr_AI)):
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
                    bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + self.CANAddr_AI),
                                                                           self.m_transmitData,1)
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
                    self.item_signal.emit([1, 2, 2, '进入模式失败'])
                    self.item_signal.emit([2, 2, 2, '进入模式失败'])

        # for d in range(500): #只在循环测试中使用
        #     self.result_signal.emit(f"第{d+1}次循环"+self.HORIZONTAL_LINE)#只在循环测试中使用
        #     self.testNum = 4#只在循环测试中使用
        if self.isCalibrate:
            bool_calibrate = self.calibrateAI()
            if not bool_calibrate:
                self.pass_signal.emit(False)
                #标定出错直接取消后续的测试和表格生成
                self.isTest= False
                self.isExcel=False
        else:
            bool_calibrate = False

        if self.isTest:
            # self.isAICurPass = True#只在循环测试中使用
            # self.isAIVolPass = True#只在循环测试中使用
            self.initPara_array()
            if self.isAITestVol:
                bool_testVol = self.testAI('AIVoltage')
                if not bool_testVol:
                    self.isExcel = False
                    self.pass_signal.emit(False)
            else:
                bool_testVol = False

            if self.isAITestCur:
                bool_testCur = self.testAI('AICurrent')
                if not bool_testCur:
                    self.isExcel = False
                    self.pass_signal.emit(False)
            else:
                bool_testCur = False
        else:
            bool_testVol = False
            bool_testCur = False

        if self.isExcel:
            self.result_signal.emit('开始生成校准校验表…………' + self.HORIZONTAL_LINE)
            code_array = [self.module_pn, self.module_sn, self.module_rev]
            station_array = [self.isAIPassTest, self.isAITestVol, self.isAITestCur]
            excel_bool, book, sheet, self.AI_row= otherOption.generateExcel(code_array,

                                                                            station_array,self.AI_Channels, 'AI')
            if not excel_bool:
                self.result_signal.emit('校准校验表生成出错！请检查代码！' + self.HORIZONTAL_LINE)
            else:
                self.fillInAIData(self.isAIPassTest, book, sheet)
                self.result_signal.emit('生成校准校验表成功' + self.HORIZONTAL_LINE)

            if not self.isAIVolPass or not self.isAICurPass:
                self.result_signal.emit(f'不通过原因：\n{self.errorInf}' + self.HORIZONTAL_LINE)
        elif not self.isExcel:
            self.result_signal.emit('测试停止，未生成校准校验表！' + self.HORIZONTAL_LINE)

        self.allFinished_signal.emit()
        self.pass_signal.emit(True)
        # if bool_calibrate:
        #     self.pass_signal.emit(True)
        # elif not self.isCalibrate and (bool_testVol or bool_testCur):
        #     self.pass_signal.emit(True)
        # elif not self.isCalibrate and not self.isTest:
        #     self.pass_signal.emit(True)
        # elif not self.isCalibrate and (bool_testVol or bool_testCur):


    # def CAN_init(self):
    #     CAN_option.close(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX)
    #     time.sleep(0.1)
    #     QApplication.processEvents()
    #
    #     CAN_option.connect(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX, CAN_option.CAN_INDEX)
    #
    #     CAN_option.init(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX, CAN_option.CAN_INDEX, init_config)
    #
    #     CAN_option.start(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX, CAN_option.CAN_INDEX)

    def testAI(self,testType):
        #1.退出标定模式
        self.setAIChOutCalibrate()
        #2.判断测试类型
        if testType == 'AIVoltage':
            self.pauseOption()
            if not self.is_running:
                return False
            # 每次切换电压电流量程初始化CAN分析仪
            # self.CAN_init()
            self.result_signal.emit('切换到电压模式' + self.HORIZONTAL_LINE)
            #CAN_option.close(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX)
            #self.can_start()
            time.sleep(1)
            self.testNum = self.testNum - 1
            self.m_transmitData = [0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]

            bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x200 + self.CANAddr_relay, self.m_transmitData,1)
            if not bool_transmit:
                self.result_signal.emit('继电器1切换错误，请停止检查设备！' + self.HORIZONTAL_LINE)
                # print('继电器1切换错误，请停止检查设备！')
                return False

            bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x200 + self.CANAddr_relay + 1, self.m_transmitData,1)
            if not bool_transmit:
                self.result_signal.emit('继电器2切换错误，请停止检查设备！' + self.HORIZONTAL_LINE)
                # print('继电器2切换错误，请停止检查设备！')
                return False
            time.sleep(0.3)
            #3-1.测试程序
            if not self.AITestLoop(testType):
                return False
            self.m_transmitData = [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]

            bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x200 + self.CANAddr_relay, self.m_transmitData,1)
            if not bool_transmit:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit('继电器1切换错误，请停止检查设备！' + self.HORIZONTAL_LINE)
                # print('继电器1切换错误，请停止检查设备！')
                return False

            bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x200 + self.CANAddr_relay + 1,
                                                                   self.m_transmitData,1)
            if not bool_transmit:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit('继电器2切换错误，请停止检查设备！' + self.HORIZONTAL_LINE)
                # print('继电器2切换错误，请停止检查设备！')
                return False
            time.sleep(0.3)

        if testType == 'AICurrent':
            self.pauseOption()
            if not self.is_running:
                return False
            # 每次切换电压电流量程初始化CAN分析仪
            # self.CAN_init()
            self.result_signal.emit('切换到电流模式'+self.HORIZONTAL_LINE)

            time.sleep(1)
            self.testNum = self.testNum - 1
            self.m_transmitData = [0x06, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]

            bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x200 + self.CANAddr_relay,
                                                                   self.m_transmitData,1)
            if not bool_transmit:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit('继电器1切换错误，请停止检查设备！' + self.HORIZONTAL_LINE)
                # print('继电器1切换错误，请停止检查设备！')
                return False

            bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x200 + self.CANAddr_relay + 1,
                                                                   self.m_transmitData,1)
            if not bool_transmit:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit('继电器2切换错误，请停止检查设备！' + self.HORIZONTAL_LINE)
                # print('继电器2切换错误，请停止检查设备！')
                return False
            time.sleep(0.3)

            if not self.AITestLoop(testType):
                return False

            self.m_transmitData = [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x200 + self.CANAddr_relay, self.m_transmitData,1)
            if not bool_transmit:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit('继电器切换错误，请停止检查设备！' + self.HORIZONTAL_LINE)
                # print('继电器切换错误，请停止检查设备！')
                return False

            bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x200 + self.CANAddr_relay + 1,
                                                                   self.m_transmitData,1)
            if not bool_transmit:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit('继电器2切换错误，请停止检查设备！' + self.HORIZONTAL_LINE)
                # print('继电器2切换错误，请停止检查设备！')
                return False
            time.sleep(0.3)

        return True

    def AITestLoop(self,type:str):

        testStart_time = time.time()
        m_valueTheory = []
        m_arrayVal = []
        vol_cur_testNum = 0 #电压/电流的量程数
        m_name=''
        # isPass = True
        if type == 'AIVoltage':
            vol_cur_testNum = 5 #电压有5种量程
            self.volValue_array=[self.voltageTheory_1010,self.voltageTheory_0505,self.voltageTheory_0005,
                            self.voltageTheory_0010,self.voltageTheory_0105]
            self.volName_array=[self.arrayVol_1010,self.arrayVol_0505,self.arrayVol_0005,
                            self.arrayVol_0010,self.arrayVol_0105]
            self.volRange_array=[27648*2,27648*2,27648,27648,27648]
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit('AI模块电压测试开始......' + self.HORIZONTAL_LINE)
            # print('AI模块电压测试开始......' + self.HORIZONTAL_LINE)
            self.item_signal.emit([7, 1, 0, ''])
            if not self.channelZero():
                return False

            for typeNum in range(vol_cur_testNum): #对每个量程进行测试
                m_name = self.vol_name_array[typeNum]
                m_valueTheory = self.volValue_array[typeNum]
                m_arrayVal = self.volName_array[typeNum]
                m_range = self.volRange_array[typeNum]
                # #修改AI量程
                for i in range(self.m_Channels):
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    if not self.setAIInputType(i + 1, type, typeNum):
                        return False
                # 修改AO量程
                for i in range(self.m_Channels):
                    if not self.setAOInputType(i + 1, type, typeNum):
                        return False

                if not self.vol_cur_test(type,m_name,m_valueTheory, m_arrayVal, m_range,typeNum):
                    self.isAIVolPass = False
        elif type == 'AICurrent':
            vol_cur_testNum = 2 #电流有2种量程
            self.curValue_array = [self.currentTheory_0420, self.currentTheory_0020]
            self.curName_array = [self.arrayCur_0420, self.arrayCur_0020]
            self.curRange_array = [27648, 27648]

            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit('AI模块电流测试开始......' + self.HORIZONTAL_LINE)
            # print('AI模块电流测试开始......' + self.HORIZONTAL_LINE)
            self.item_signal.emit([8, 1, 0, ''])
            if not self.channelZero():
                return False

            for typeNum in range(vol_cur_testNum):  # 对每个量程进行测试
                m_name = self.cur_name_array[typeNum]
                m_valueTheory = self.curValue_array[typeNum]
                m_arrayVal = self.curName_array[typeNum]
                m_range =self.curRange_array[typeNum]
                # #修改AI量程
                for i in range(self.m_Channels):
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    if not self.setAIInputType(i + 1, type, typeNum):
                        return False
                # 修改AO量程
                for i in range(self.m_Channels):
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    if not self.setAOInputType(i + 1, type, typeNum):
                        return False
                if not self.vol_cur_test(type,m_name, m_valueTheory, m_arrayVal, m_range, typeNum):
                    self.isAICurPass = False

        # 数据清零
        if not self.normal_writeValuetoAO(0):
            return False
        self.pauseOption()
        if not self.is_running:
            return False
        self.result_signal.emit('测试结束' + self.HORIZONTAL_LINE)

        testEnd_time = time.time()
        testTest_time = round(testEnd_time - testStart_time,2)
        if self.isAIVolPass and type == 'AIVoltage':
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f'电压通过：{self.isAIVolPass}\n\n')
            self.item_signal.emit([7, 2, 1, f'{testTest_time}'])
        elif not self.isAIVolPass and type == 'AIVoltage':
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f'电压不通过：{self.isAIVolPass}\n\n')
            self.item_signal.emit([7, 2, 2, f'{testTest_time}'])
        if self.isAICurPass and type == 'AICurrent':
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f'电流通过：{self.isAICurPass}\n\n')
            self.item_signal.emit([8, 2, 1, f'{testTest_time}'])
        elif not self.isAICurPass and type == 'AICurrent':
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f'电流不通过：{self.isAICurPass}\n\n')
            self.item_signal.emit([8, 2, 2, f'{testTest_time}'])
        self.isAIPassTest = self.isAIVolPass & self.isAICurPass
        # return self.isAIPassTest
        return True


    def vol_cur_test(self,type:str,m_name,m_valueTheory,m_arrayVal,m_range,typeNum:int)->bool:
        """
        单个量程测试
        :param type: str
        :param m_valueTheory: list
        :param m_arrayVal: list
        :param m_range:list
        :param typeNum: int
        :return: bool
        """
        for i in range(5):
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f'{m_arrayVal[i]}\n\n')
            # print(f'{m_arrayVal[i]}\n\n')
            # for n in range(10):
            if not self.normal_writeValuetoAO(m_valueTheory[i]):
                return False
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f'发送的数据：{m_valueTheory[i]}')
            # print(f'{m_valueTheory[i]}\n\n')

            self.result_signal.emit(f'等待信号稳定……\n')
            # tt=time.time()
            # while True:
            #     if time.time()-tt >= int(self.waiting_time/1000):
            #         break
            #     waitFlag, rece_wait = self.receiveAIData()
            #     if waitFlag == 'stopReceive':
            #         return False
            # recv = [0, 0, 0, 0]
            # if type == 'AIVoltage':
            #     self.error_value = 422
                # self.error_value = self.vol_error_value[3]
                # if int(m_valueTheory[i]) == 27648:
                #     self.standardValue = self.vol_high_standardValue_array[typeNum]
                # elif int(m_valueTheory[i]) == 0 or int(m_valueTheory[i]) == -27648:
                #     self.standardValue = self.vol_low_standardValue_array[typeNum]

            # elif type == 'AICurrent':
            #     self.error_value = 636
                # self.error_value = self.cur_high_error_value[0]
                # if int(m_valueTheory[i]) == 27648:
                #     self.error_value = self.cur_high_error_value[typeNum]
                #     self.standardValue = self.cur_high_standardValue_array[typeNum]
                # elif int(m_valueTheory[i]) == 0:
                #     self.error_value = self.cur_low_error_value[typeNum]
                #     self.standardValue = self.cur_low_standardValue_array[typeNum]
            t_wait =time.time()
            warning_sign = False
            while True:
                if (time.time()-t_wait)*1000 >= self.waiting_time:
                    self.result_signal.emit(f'规定时间内无法接收到稳定信号，请检查通道是否损坏！\n')
                    self.result_signal.emit(f'各通道实际接收数据：{rece_wait[0]}、{rece_wait[1]}、'
                                            f'{rece_wait[2]}、{rece_wait[3]}\n')
                    self.result_signal.emit(f'最大差值：{abs(int(m_valueTheory[0]/1000))}')
                    self.result_signal.emit(f'各通道实际差值：{abs(rece_wait[0] - m_valueTheory[i])}、{abs(rece_wait[1] - m_valueTheory[i])}、'
                                            f'{abs(rece_wait[2] - m_valueTheory[i])}、{abs(rece_wait[3] - m_valueTheory[i])}\n')
                    self.result_signal.emit(f'量程”{m_name}“的 {m_arrayVal[i]}不通过\n')
                    warning_sign = True
                    break
                waitFlag, rece_wait = self.receiveAIData()
                # print(f"rece_wait:{rece_wait}")
                # print(f"m_valueTheory[i]:{m_valueTheory[i]}")
                if waitFlag == 'stopReceive':
                    return False
                if m_valueTheory[i]==0:
                    if (abs(rece_wait[0] - m_valueTheory[i]) <= 28 and
                        abs(rece_wait[1] - m_valueTheory[i]) <= 28 and
                        abs(rece_wait[2] - m_valueTheory[i]) <= 28 and
                        abs(rece_wait[3] - m_valueTheory[i]) <= 28):
                        break
                else:
                    if (abs(rece_wait[0] - m_valueTheory[i]) <= abs(int(m_valueTheory[0]/1000)) and
                        abs(rece_wait[1] - m_valueTheory[i]) <= abs(int(m_valueTheory[0]/1000)) and
                        abs(rece_wait[2] - m_valueTheory[i]) <= abs(int(m_valueTheory[0]/1000)) and
                        abs(rece_wait[3] - m_valueTheory[i]) <= abs(int(m_valueTheory[0]/1000))):
                        break
                time.sleep(0.2)
            if warning_sign:
                return False

            # while True:
            #     if (time.time() - t_wait)*1000 >= self.waiting_time:
            #         self.result_signal.emit(f'规定时间内无法接收到稳定信号，请检查通道是否损坏！\n')
            #         warning_sign=True
            #         break
            #     # self.receiveAIData()
            #     bool_receive, self.m_can_obj = CAN_option.receiveCANbyID(0x280 + self.CANAddr_AI, self.waiting_time)
            #     for j in range(self.m_Channels):
            #         data = bytes([self.m_can_obj.Data[j * 2], self.m_can_obj.Data[j * 2 + 1]])
            #         # recv[i] = value = struct.unpack('<h', data)[0]
            #         recv[j] = struct.unpack('<h', data)[0]
            #     if ((recv[0] - m_valueTheory[i]) < self.error_value and (recv[1] - m_valueTheory[i]) < self.error_value
            #             and (recv[2] - m_valueTheory[i]) < self.error_value and (recv[3] - m_valueTheory[i]) < self.error_value):
            #         break
            # if warning_sign:
            #     return False



            time1 = time.time()
            reReceiveNum = 0

            while True:
                if (time.time() - time1)*1000 > self.waiting_time:
                    self.messageBox_signal.emit(['警告', '数据接收超时！'])
                    break

                bool_receiveAI,usReceValue = self.receiveAIData()
                if not bool_receiveAI:
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    self.messageBox_signal.emit(['警告', '数据接收出错，请检查设备'])
                    return False
                elif bool_receiveAI == 'stopReceive':
                    return False
                else:
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    self.result_signal.emit('数据接收成功!\n\n')
                    # print(f'time:{time.time()}\nusReceValue={usReceValue}\nm_valueTheory[i]={m_valueTheory[i]}\n\n')
                    break
            # 各通道的精度数组
            chPrecision = [0 for x in range(self.m_Channels)]
            #计算4个通道的精度
            for j in range(self.m_Channels):
                # self.isPause()
                # if not self.isStop():
                #     return
                self.pauseOption()
                if not self.is_running:
                    return False
                #取量程中的5个点进行精度计算，正向量程的第5个点和双向量程的第3个点为0，若有一点误差收到值有可能会为在65535附近
                if (i == 4 and abs(usReceValue[j] - 65535) < 100) or (i == 2 and abs(usReceValue[j] - 65535) < 100):
                    fPrecision = round(self.GetPrecision(usReceValue[j],65535, m_range), 2)
                    #将接收到的值转换到0附近进行显示
                    usReceValue[j] = (usReceValue[j] - 65535) -1
                else:
                    fPrecision = round(self.GetPrecision(usReceValue[j], m_valueTheory[i], m_range), 2)
                self.result_signal.emit(f'\n接收到AO通道 {j + 1} 数据:{usReceValue[j]}\n\n')
                # print(f'\n接收到AO通道 {j + 1} 数据:{usReceValue[j]}\n\n')
                chPrecision[j] = fPrecision
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit(f'----------精度(‰)：{fPrecision}  ')
                # print(f'----------精度(‰)：{fPrecision}')
                if abs(fPrecision) < 1:
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    self.result_signal.emit('\t满足精度\n\n')
                    # print('满足精度\n\n')
                    if type == 'AIVoltage':
                        self.isAIVolPass &= True
                        self.pauseOption()
                        if not self.is_running:
                            return False
                        # self.result_signal.emit(f'满足精度：{self.isAIVolPass}')
                    else:
                        self.isAICurPass &= True
                        self.pauseOption()
                        if not self.is_running:
                            return False
                        # self.result_signal.emit(f'满足精度：{self.isAICurPass}')
                else:
                    # print('\t不满足精度\n\n')
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    self.result_signal.emit('不满足精度\n\n')
                    if type == 'AIVoltage':
                        self.isAIVolPass &= False
                        self.pauseOption()
                        if not self.is_running:
                            return False
                        # self.result_signal.emit(f'不满足精度：{self.isAIVolPass}')
                    else:
                        self.isAICurPass &= False
                        self.pauseOption()
                        if not self.is_running:
                            return False
                        # self.result_signal.emit(f'不满足精度：{self.isAICurPass}')
                if j == self.m_Channels - 1:
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    self.result_signal.emit(self.HORIZONTAL_LINE)
                time.sleep(0.1)
                if type == 'AIVoltage':
                    self.volReceValue[typeNum][i] = usReceValue
                    self.volPrecision[typeNum][i] = chPrecision
                elif type == 'AICurrent':
                    self.curReceValue[typeNum][i] = usReceValue
                    self.curPrecision[typeNum][i] = chPrecision
        time.sleep(0.5)
        return True
    def GetPrecision(self, receValue, theoryValue, range):
        return (receValue - theoryValue) * 1000 / range

    def  receiveAIData(self):
        can_id = 0x280 + self.CANAddr_AI
        recv = [0,0,0,0]
        time1 = time.time()
        while True:
            if (time.time() - time1)*1000 > self.waiting_time:
                return False,0
            bool_receive,self.m_can_obj = CAN_option.receiveCANbyID(can_id, self.waiting_time,1)
            # QApplication.processEvents()
            if bool_receive == 'stopReceive':
                return 'stopReceive', recv
            if bool_receive:
                break
        for i in range(self.m_Channels):
            data = bytes([self.m_can_obj.Data[i * 2], self.m_can_obj.Data[i * 2 + 1]])
            # # print(f'i= {i}')
            # recv[i] = self.m_can_obj.Data[i*2] | self.m_can_obj.Data[i*2+1] << 8
            # # print(f'recv[{i}]={recv[i]}')
            # self.isPause()
            # if not self.isStop():
            #     return
            recv[i] = value = struct.unpack('<h', data)[0]
        # print(f'recv= {recv}')
        return True,recv

    def calibrateAI(self):
        if self.isCalibrateVol == True:
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit('继电器切换到电压模式' + self.HORIZONTAL_LINE)
            #每次切换电压电流量程初始化CAN分析仪
            # self.CAN_init()
            # self.testNum = self.testNum - 1
            self.m_transmitData = [0x01,0x01,0x00,0x00,0x00,0x00,0x00,0x00]
            try:
                bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x200 + self.CANAddr_relay,
                                                                       self.m_transmitData,1)
                if not bool_transmit:
                    return False
            except:
                self.messageBox_signal.emit(['错误提示', '继电器1切换错误，请停止检查设备！'])
                # QMessageBox(QMessageBox.Critical, '错误提示', '继电器切换错误，请停止检查设备！').exec_()
                return False

            try:
                bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x200 + self.CANAddr_relay+1,
                                                                       self.m_transmitData,1)
                if not bool_transmit:
                    return False
            except:
                self.messageBox_signal.emit(['错误提示', '继电器2切换错误，请停止检查设备！'])
                # QMessageBox(QMessageBox.Critical, '错误提示', '继电器切换错误，请停止检查设备！').exec_()
                return False
            time.sleep(0.3)

            self.pauseOption()
            if not self.is_running:
                return False
            self.item_signal.emit([5, 1, 0, ''])
            # self.itemOperation(mTable, 5, 1, 0, '')
            calibrateStart_time = time.time()
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f'开始标定AI电压：' + self.HORIZONTAL_LINE)
            for typeNum in range(5):
                bool_calibrate = self.calibrateAI_vol_cur('AIVoltage',typeNum)
                if not bool_calibrate:
                    calibrateFailEnd_time = time.time()
                    calibrateFail_time = round(calibrateFailEnd_time - calibrateStart_time, 1)
                    self.item_signal.emit([5, 2, 2, f'{calibrateFail_time}'])
                    return False
                self.pauseOption()
                if not self.is_running:
                    return False
            calibrateEnd_time = time.time()
            calibrateTest_time = round(calibrateEnd_time - calibrateStart_time, 2)
            self.pauseOption()
            if not self.is_running:
                return False
            if bool_calibrate:
                self.item_signal.emit([5, 2, 1,f'{calibrateTest_time}'])
            else:
                self.item_signal.emit([5, 2, 2,f'{calibrateTest_time}'])
            # self.itemOperation(mTable, 5, 2, 1, f'{calibrateTest_time}')
            self.m_transmitData = [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
            # time.sleep(0.5)
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x200 + self.CANAddr_relay, self.m_transmitData,1)
            if not bool_transmit:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit('继电器1切换错误，请停止检查设备！' + self.HORIZONTAL_LINE)
                # print('继电器1切换错误，请停止检查设备！')
                # self.work_thread.stopFlag.isSet()
                # self.isStop()
                return False

            bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x200 + self.CANAddr_relay+1, self.m_transmitData,1)
            if not bool_transmit:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit('继电器2切换错误，请停止检查设备！' + self.HORIZONTAL_LINE)
                # print('继电器2切换错误，请停止检查设备！')
                # self.work_thread.stopFlag.isSet()
                # self.isStop()
                return False
            time.sleep(0.3)

        elif not self.isCalibrateVol:
            self.pauseOption()
            if not self.is_running:
                return False
            self.item_signal.emit([5,0,0,''])

        if self.isCalibrateCur == True:
            self.pauseOption()
            if not self.is_running:
                return False
            # 每次切换电压电流量程初始化CAN分析仪
            # self.CAN_init()
            self.result_signal.emit('继电器切换到电流模式' + self.HORIZONTAL_LINE)

            # self.testNum = self.testNum - 1
            self.m_transmitData = [0x06, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
            try:
                bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x200 + self.CANAddr_relay,
                                                                       self.m_transmitData,1)
                if not bool_transmit:
                    self.pauseOption()
                    if not self.is_running:
                        return False
            except:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.messageBox_signal.emit(['错误提示', '继电器1切换错误，请停止检查设备！'])
                # self.result_signal.emit('继电器1切换错误，请停止检查设备！' + self.HORIZONTAL_LINE)
                return False
            try:
                bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x200 + self.CANAddr_relay + 1,
                                                                       self.m_transmitData,1)
                if not bool_transmit:
                    self.pauseOption()
                    if not self.is_running:
                        return False
            except:
                self.pauseOption()
                if not self.is_running:
                    return False
                # self.messageBox_signal.emit(['错误提示', '继电器2切换错误，请停止检查设备！'])
                self.result_signal.emit('继电器2切换错误，请停止检查设备！' + self.HORIZONTAL_LINE)
                return False
            time.sleep(0.3)

            # self.itemOperation(mTable, 6, 1, 0, '')
            self.pauseOption()
            if not self.is_running:
                return False
            self.item_signal.emit([6, 1, 0, ''])
            calibrateStart_time = time.time()
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f'开始标定AI电流：' + self.HORIZONTAL_LINE)
            for typeNum in range(2):
                bool_calibrate = self.calibrateAI_vol_cur('AICurrent',typeNum)
                if not bool_calibrate:
                    return False
                self.pauseOption()
                if not self.is_running:
                    return False
            calibrateEnd_time = time.time()
            calibrateTest_time = round(calibrateEnd_time - calibrateStart_time, 2)
            # self.itemOperation(mTable, 6, 2, 1, f'{calibrateTest_time}')
            self.pauseOption()
            if not self.is_running:
                return False
            self.item_signal.emit([6, 2, 1, f'{calibrateTest_time}'])
            #CAN_option.close(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX)
            #self.can_start()
            # time.sleep(1)
            self.m_transmitData = [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]

            bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x200 + self.CANAddr_relay, self.m_transmitData,1)
            if not bool_transmit:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit('继电器1切换错误，请停止检查设备！' + self.HORIZONTAL_LINE)
                # print('继电器1切换错误，请停止检查设备！')
                return False

            bool_transmit, self.m_can_obj = CAN_option.transmitCAN(0x200 + self.CANAddr_relay+1, self.m_transmitData,1)
            if not bool_transmit:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit('继电器2切换错误，请停止检查设备！' + self.HORIZONTAL_LINE)
                # print('继电器2切换错误，请停止检查设备！')
                return False
            time.sleep(0.3)

        elif not self.isCalibrateCur:
            self.pauseOption()
            if not self.is_running:
                return False
            self.item_signal.emit([6, 0, 0, ''])

        return True

    def calibrateAI_vol_cur(self,type:str,typeNum:int):
        # [(367422 - 156866), (314993 - 209295), (314679 - 262038), (367212 - 261933),(314658 - 272545)]
        #                    -/+10V   -/+5V   0-5V   0-10V   1-5V
        self.vol_maxRange_array = [210556, 105698, 52641, 105279, 42113]
        # cur_maxRange_array = [(315094 - 272538), (315115 - 262037)]
        #                    4-20mA  0-20mA
        self.cur_maxRange_array = [42556, 53078]

        self.vol_highValue_array = [0x6c00,0x6c00,0x6c00,0x6c00,0x6c00]
        self.vol_lowValue_array = [-27648,-27648,0,0,0]
        self.cur_highValue_array = [0x6c00,0x6c00]
        self.cur_lowValue_array = [0,0]

        maxRange = 0
        lowValue = 0
        highValue = 0
        if type == 'AIVoltage':
            highValue = self.vol_highValue_array[typeNum]
            lowValue = self.vol_lowValue_array[typeNum]
            maxRange = self.vol_maxRange_array[typeNum] #上限最大值- 下限最小值
        if type == 'AICurrent':
            highValue = self.cur_highValue_array[typeNum]
            lowValue = self.cur_lowValue_array[typeNum]
            maxRange = self.cur_maxRange_array[typeNum]  # 上限最大值- 下限最小值

        #self.CAN_init()
        #0.防止之前可能在标定模式中，先退出标定模式
        self.setAIChOutCalibrate()

        """1.通道归零"""
        self.channelZero()

        """2.初始化模式"""
        self.setAIChOutCalibrate() #退出标定模式
        time.sleep(0.1)

        """3.设置AI量程"""
        for i in range(self.m_Channels):
            channel = i + 1
            if not self.setAIInputType(channel, type,typeNum):
                return False

        """4.设置AO量程"""
        for i in range(self.m_Channels):
            channel = i + 1
            if not self.setAOInputType(channel, type,typeNum):
                return False

        """5.进入标定模式"""
        self.pauseOption()
        if not self.is_running:
            return False
        self.result_signal.emit(f'进入标定模式\n'+self.HORIZONTAL_LINE)
        if not self.setAIChInCalibrate():
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit('进入标定模式失败，测试结束。' + self.HORIZONTAL_LINE)
            return False

        """6.向AO模块写输出值，并接收配套AI模块的测量值"""
        channelNum = 1
        if self.oneOrAll == 2:#2:标定所有通道
            channelNum = self.m_Channels
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(self.HORIZONTAL_LINE + '标定所有通道' + self.HORIZONTAL_LINE)
            # print(self.HORIZONTAL_LINE + '标定所有通道' + self.HORIZONTAL_LINE)
        if channelNum == 1:
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(self.HORIZONTAL_LINE + '仅标定第一个通道' + self.HORIZONTAL_LINE)
            # print(self.HORIZONTAL_LINE + '仅标定第一个通道' + self.HORIZONTAL_LINE)
        bool_receive,usArrayHigh = self.AI_receive_WriteToAO(highValue,type,channelNum,typeNum)
        if not bool_receive:
            self.pauseOption()
            if not self.is_running:
                return False
            self.pauseOption()
            self.result_signal.emit('AI模块标定结束！' + self.HORIZONTAL_LINE)
            # print(self.HORIZONTAL_LINE + 'AI模块标定结束！' + self.HORIZONTAL_LINE)


            """.8退出标定模式"""
            self.setAIChOutCalibrate()
            self.pauseOption()
            if not self.is_running:
                return False
            self.pauseOption()
            self.result_signal.emit('退出标定模式！' + self.HORIZONTAL_LINE)
            # print('退出标定模式！' + self.HORIZONTAL_LINE)

            """9.通道归零"""
            self.channelZero()

            # 后续测试取消
            self.isTest=False
            # 这里return False会导致软件界面卡死，暂时不知道如何解决
            #2023.11.13已解决界面卡死问题
            return False
        if bool_receive == 'stopReceive':
            return False
        self.pauseOption()
        if not self.is_running:
            return False
        self.result_signal.emit(f"等待1秒……\n\n")
        time.sleep(1)
        bool_receive,usArrayLow = self.AI_receive_WriteToAO(lowValue,type,channelNum,typeNum)
        if not bool_receive:
            self.pauseOption()
            if not self.is_running:
                return False
            self.pauseOption()
            self.result_signal.emit('AI模块标定结束！' + self.HORIZONTAL_LINE)
            # print(self.HORIZONTAL_LINE + 'AI模块标定结束！' + self.HORIZONTAL_LINE)

            """.8退出标定模式"""
            self.setAIChOutCalibrate()
            self.pauseOption()
            if not self.is_running:
                return False
            self.pauseOption()
            self.result_signal.emit('退出标定模式！' + self.HORIZONTAL_LINE)
            # print('退出标定模式！' + self.HORIZONTAL_LINE)

            """9.通道归零"""
            self.channelZero()

            # 后续测试取消
            self.isTest = False
            #这里return False会导致软件界面卡死，暂时不知道如何解决
            # 2023.11.13已解决界面卡死问题
            return False
        for i in range(channelNum):
            # self.isPause()
            # if not self.isStop():
            #     return
            if abs(usArrayHigh[i] - usArrayLow[i]) < 10 or abs(usArrayHigh[i] - usArrayLow[i] > maxRange):
                self.messageBox_signal.emit(['警告', f'请检查通道{i+1}接线或者模块是否存在问题,并重新开始标定！'])
                # reply = QMessageBox.warning(None, '警告', '请检查接线或者模块是否存在问题,并重新开始标定！',
                #                             QMessageBox.AcceptRole | QMessageBox.RejectRole, QMessageBox.AcceptRole)
                #退出标定模式并进行通道归零
                self.setAIChOutCalibrate()
                self.pauseOption()
                if not self.is_running:
                    return False
                self.pauseOption()
                self.result_signal.emit('退出标定模式！' + self.HORIZONTAL_LINE)
                # print('退出标定模式！' + self.HORIZONTAL_LINE)

                # if typeNum == 4:
                """通道归零"""
                self.channelZero()
                return False
        """7.计算并更新零值和量程值"""
        self.pauseOption()
        if not self.is_running:
            return False
        self.result_signal.emit(self.HORIZONTAL_LINE+'计算AI模块零值和量程值\n\n')
        for i in range(channelNum):
            # self.isPause()
            # if not self.isStop():
            #     return
            usSpanValue = usArrayHigh[i]
            usZeroValue =usArrayLow[i]
            # usSpanValue = int(self.calcSpan(usArrayHigh[i],usArrayLow[i],highValue,lowValue))
            # usZeroValue = int(self.calcZero(usArrayHigh[i], usArrayLow[i], highValue, lowValue,type))
            # print(f'{i+1}.计算得到AI通道{i+1} 零值：{usZeroValue}，量程值：{usSpanValue}')
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f'{i+1}.计算得到AI通道{i+1} 零值：{usZeroValue}，量程值：{usSpanValue}')

            """向通道写入零值与量程值"""
            if not self.writeParaToChannel(i+1, usZeroValue, usSpanValue):
                self.pauseOption()
                if not self.is_running:
                    return False
                if type == 'AIVoltage':
                    self.result_signal.emit(f'更新通道{i+1}在"{self.vol_name_array[typeNum]}"量程的标定值失败！\n\n')
                elif type == 'AICurrent':
                    self.result_signal.emit(f'更新通道{i + 1}在"{self.cur_name_array[typeNum]}"量程的标定值失败！\n\n')
            else:
                self.pauseOption()
                if not self.is_running:
                    return False
                if type == 'AIVoltage':
                    self.result_signal.emit(f'成功更新通道{i + 1}在"{self.vol_name_array[typeNum]}"量程的标定值！\n\n')
                elif type == 'AICurrent':
                    self.result_signal.emit(f'成功更新通道{i + 1}在"{self.cur_name_array[typeNum]}"量程的标定值！\n\n')
            #若只标定第一个通道，则需要把第一个通道的零值和量程值也写入其他三个通道
            if channelNum == 1:
                self.pauseOption()
                if not self.is_running:
                    return False
                self.result_signal.emit('更新其他三个通道的标定值')
                # print('更新其他三个通道的标定值')
                self.writeParaToChannel(i + 2, usZeroValue, usSpanValue)
                self.writeParaToChannel(i + 3, usZeroValue, usSpanValue)
                self.writeParaToChannel(i + 4, usZeroValue, usSpanValue)

        self.pauseOption()
        if not self.is_running:
            return False
        self.pauseOption()
        self.result_signal.emit('AI模块标定结束！' + self.HORIZONTAL_LINE)
        # print(self.HORIZONTAL_LINE+'AI模块标定结束！' + self.HORIZONTAL_LINE)

        """.8退出标定模式"""
        self.setAIChOutCalibrate()
        self.pauseOption()
        if not self.is_running:
            return False
        self.pauseOption()
        self.result_signal.emit('退出标定模式！' + self.HORIZONTAL_LINE)
        # print('退出标定模式！' + self.HORIZONTAL_LINE)

        if typeNum == 4:
            """9.通道归零"""
            self.channelZero()
        return True


    # 向通道写入标定参数
    def writeParaToChannel(self, AIChannel, usZero, usSpan):
        can_id = 0x600 + self.CANAddr_AI
        self.m_transmitData[0] = 0x27
        self.m_transmitData[2] = 0x64
        self.m_transmitData[3] = AIChannel
        self.m_transmitData[7] = 0x00

        # AI保存校正下限值
        self.m_transmitData[1] = 0x3d
        self.m_transmitData[4] = (usZero & 0xff)
        self.m_transmitData[5] = ((usZero >> 8) & 0xff)
        self.m_transmitData[6] = ((usZero >> 16) & 0xff)
        bool_transmit_low, self.m_can_obj = CAN_option.transmitCAN(can_id, self.m_transmitData,1)

        time.sleep(0.1)

        # AI保存校正上限值
        self.m_transmitData[1] = 0x3e
        self.m_transmitData[4] = (usSpan & 0xff)
        self.m_transmitData[5] = ((usSpan >> 8) & 0xff)
        self.m_transmitData[6] = ((usSpan >> 16) & 0xff)
        bool_transmit_high, self.m_can_obj = CAN_option.transmitCAN(can_id, self.m_transmitData,1)
        return bool_transmit_low&bool_transmit_high

    def AI_receive_WriteToAO(self,value,type,channelNum,typeNum):
        self.vol_high_standardValue_array = [367002, 314573, 314573, 367002, 314573]
        self.cur_high_standardValue_array = [314777, 314777]
        self.vol_low_standardValue_array = [157286, 209715, 262144, 262144, 272630]
        self.cur_low_standardValue_array = [272671, 262144]
        self.vol_error_value = [841, 841, 212, 422, 170]
        self.cur_high_error_value = [636, 678]
        self.cur_low_error_value = [265, 214]
        # # print('写入数据并等待信号稳定………' + self.HORIZONTAL_LINE)
        if type == 'AIVoltage':
            self.error_value=self.vol_error_value[typeNum]
            if int(value) == 27648:
                self.standardValue = self.vol_high_standardValue_array[typeNum]
            elif int(value) == 0 or int(value) == -27648:
                self.standardValue = self.vol_low_standardValue_array[typeNum]
            # headInf = self.arrayVol_1010
        elif type == 'AICurrent':
            if int(value) == 27648:
                self.error_value =self.cur_high_error_value[typeNum]
                self.standardValue = self.cur_high_standardValue_array[typeNum]
            elif int(value) == 0:
                self.error_value = self.cur_low_error_value[typeNum]
                self.standardValue = self.cur_low_standardValue_array[typeNum]

        self.pauseOption()
        if not self.is_running:
            return False,0
        if value >0:
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f'写入第一个数据点（{self.standardValue}）并等待信号稳定………' + self.HORIZONTAL_LINE)
        else:
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f'写入第二个数据点（{self.standardValue}）并等待信号稳定………' + self.HORIZONTAL_LINE)
        if not self.normal_writeValuetoAO(value):
            return False

        if channelNum == 1:
            usRecValue = [0]
        elif channelNum == 4:
            usRecValue = [0, 0, 0, 0]

        inf = '接收到的AI数据：'
        inf_average = '接收到AI数据的平均值：'
        valReceive_num = self.receive_num
        warning_flag = False

        #需要接收self.receive_num次数据后算平均值
        for i in range(self.receive_num):
            time2 = time.time()
            #一开始直接等 self.waiting_time ms，等待信号稳定再接收报文
            if i == 0:
                while True:
                    if (time.time() - time2)*1000 >= self.waiting_time:
                        self.pauseOption()
                        if not self.is_running:
                            return False,0
                        break

            bool_caReceive, usTmpValue = self.calibrate_receiveAIData(channelNum)
            if not bool_caReceive:
                self.pauseOption()
                if not self.is_running:
                    return False, 0
                self.result_signal.emit(f'{i + 1}.第{i + 1}次数据未接收到！排除该次数据！\n\n')
                valReceive_num = valReceive_num -1
                continue
            if bool_caReceive == 'stopReceive':
                return False,0
            if channelNum == 4:
                self.pauseOption()
                if not self.is_running:
                    return False, 0
                if (abs(usTmpValue[0]-self.standardValue)>200 or abs(usTmpValue[1]-self.standardValue)>200
                        or abs(usTmpValue[2]-self.standardValue)>400 or abs(usTmpValue[3]-self.standardValue)>200):
                    self.result_signal.emit(f'{i + 1}.第{i + 1}次数据：{usTmpValue}，误差过大！排除该次数据！\n\n')
                    valReceive_num = valReceive_num - 1
                    continue
            if channelNum == 1:
                self.pauseOption()
                if not self.is_running:
                    return False, 0
                if (abs(usTmpValue[0]-self.standardValue)>200):
                    self.result_signal.emit(f'{i + 1}.第{i + 1}次数据：{usTmpValue}，误差过大！排除该次数据！\n\n')
                    valReceive_num = valReceive_num - 1
                    continue
            self.pauseOption()
            if not self.is_running:
                return False, 0
            self.result_signal.emit(f'{i + 1}.第{i + 1}次{inf}{usTmpValue}  \n\n')
            for j in range(channelNum):
                usRecValue[j]=usRecValue[j]+usTmpValue[j]
            #     if i == self.receive_num - 1:
            #         usRecValue[j] = int(usRecValue[j]/valReceive_num)
            # if i ==  self.receive_num - 1:
            #     self.result_signal.emit(inf_average + f'{usRecValue}\n\n')
        if valReceive_num > 0:
            for j in range(channelNum):
                usRecValue[j] = int(usRecValue[j] / valReceive_num)
            self.result_signal.emit(inf_average + f'{usRecValue}\n\n')
        elif valReceive_num == 0:
            self.pauseOption()
            if not self.is_running:
                return False, 0
            self.result_signal.emit(f'!!!!!!!!{self.receive_num}次均未接受到正常数据，请检查接线和设备!!!!!!!!\n\n')
            return False, usRecValue
        return True,usRecValue

    def calibrate_receiveAIData(self, channelNum):
        # CAN_option.close(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX)
        # self.can_start()
        recv = [0, 0, 0, 0]
        if channelNum == 1:
            recv = [0]
        can_id = 0x600 + self.CANAddr_AI
        for i in range(channelNum):
            self.m_transmitData[0] = 0x40
            self.m_transmitData[1] = 0x3c
            self.m_transmitData[2] = 0x64
            self.m_transmitData[3] = i + 1
            self.m_transmitData[4] = 0x00
            self.m_transmitData[5] = 0x00
            self.m_transmitData[6] = 0x00
            self.m_transmitData[7] = 0x00
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN(can_id, self.m_transmitData,1)
            # self.isPause()
            # if not self.isStop():
            #     return
            while True:
                # self.isPause()
                # if not self.isStop():
                #                  return
                bool_receive, self.m_can_obj = CAN_option.receiveCANbyID(0x580 + self.CANAddr_AI, self.waiting_time,1)
                # QApplication.processEvents()
                if bool_receive:
                    break
                if bool_receive == 'stopReceive':
                    return 'stopReceive',recv
                elif not bool_receive:
                    return False,recv
            recv[i] = ((self.m_can_obj.Data[7] << 24 | self.m_can_obj.Data[6] << 16) | self.m_can_obj.Data[5] << 8) | \
                      self.m_can_obj.Data[4]

        return True,recv

    def setAIChInCalibrate(self):
        bool_all = True
        self.m_transmitData[0] = 0x23
        self.m_transmitData[1] = 0x3f
        self.m_transmitData[2] = 0x64
        self.m_transmitData[3] = 0x00
        self.m_transmitData[4] = 0x43
        self.m_transmitData[5] = 0x41
        self.m_transmitData[6] = 0x4c
        self.m_transmitData[7] = 0x49

        bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + self.CANAddr_AI), self.m_transmitData,1)
        bool_all = bool_all & bool_transmit

        self.m_transmitData[0] = 0x23
        self.m_transmitData[1] = 0x3f
        self.m_transmitData[2] = 0x64
        self.m_transmitData[3] = 0x00
        self.m_transmitData[4] = 0x53
        self.m_transmitData[5] = 0x54
        self.m_transmitData[6] = 0x41
        self.m_transmitData[7] = 0x52

        bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + self.CANAddr_AI), self.m_transmitData,1)
        bool_all = bool_all & bool_transmit

        return bool_all

    def clearList(self,array):
        for i in range(len(array)):
            array[i] = 0x00


    # 检查待检AO模块输入类型是否正确
    def setAOInputType(self, AOChannel, type:str, typeNum:int):
        if AOChannel == 1:
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f'设置AO通道量程：\n\n')

        self.m_transmitData[0] = 0x2b
        self.m_transmitData[1] = 0x10
        self.m_transmitData[2] = 0x63
        self.m_transmitData[3] = AOChannel
        self.m_transmitData[6] = 0x00
        self.m_transmitData[7] = 0x00

        self.clearList(self.m_can_obj.Data)

        if type == 'AOVoltage' or type == 'AIVoltage':
            self.m_transmitData[4] = self.vol_AORangeArray[2*typeNum]
            self.m_transmitData[5] = self.vol_AORangeArray[2*typeNum+1]
            while True:
                # self.isPause()
                # if not self.isStop():
                #                  return
                # QApplication.processEvents()
                if CAN_option.transmitCAN(0x600 + self.CANAddr_AO, self.m_transmitData,1)[0]:
                    bool_receive, self.m_can_obj = CAN_option.receiveCANbyID(0x580 + self.CANAddr_AO,
                                                                             self.waiting_time,1)
                    if bool_receive:
                        # # print(f'm_can_obj.Data[4]:{self.m_can_obj.Data[4]}')
                        # # print(f'm_can_obj.Data[5]:{self.m_can_obj.Data[5]}')
                        # # print(f'AORangeArray[0]:{self.AORangeArray[0]}')
                        # # print(f'AORangeArray[1]:{self.AORangeArray[1]}')
                        if hex(self.m_can_obj.Data[4]) == hex(self.vol_AORangeArray[2*typeNum]) and hex(
                                self.m_can_obj.Data[5]) == hex(self.vol_AORangeArray[2*typeNum+1]):
                            # print(f'{AOChannel}.成功设置AO通道{AOChannel}的量程为”{self.vol_name_array[typeNum]}“。\n\n')
                            self.pauseOption()
                            if not self.is_running:
                                return False
                            self.result_signal.emit(f'{AOChannel}.成功设置AO通道{AOChannel}的量程为”'
                                                    f'{self.vol_name_array[typeNum]}“。\n\n')
                            break
                        else:
                            self.pauseOption()
                            if not self.is_running:
                                return False
                            self.result_signal.emit(f'{AOChannel}.未成功设置AO通道{AOChannel}的量程为”'
                                                    f'{self.vol_name_array[typeNum]}“。\n\n')
                            # print(f'{AOChannel}.未成功设置AO通道{AOChannel}的量程为”{self.vol_name_array[typeNum]}“。\n\n')
                            return False
                else:
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    self.result_signal.emit(f'{AOChannel}.未成功设置AO通道{AOChannel}的量程为”'
                                            f'{self.vol_name_array[typeNum]}“。\n\n')
                    return False
                    # # print(f'{AOChannel}.未成功设置AO通道{AOChannel}的量程为”电压+/-10V“。\n\n')
        elif type == 'AOCurrent' or type == 'AICurrent':
            self.m_transmitData[4] = self.cur_AORangeArray[2*typeNum]
            self.m_transmitData[5] = self.cur_AORangeArray[2*typeNum+1]
            while True:
                # self.isPause()
                # if not self.isStop():
                #                  return
                # QApplication.processEvents()
                if CAN_option.transmitCAN(0x600 + self.CANAddr_AO, self.m_transmitData,1)[0]:
                    bool_receive, self.m_can_obj = CAN_option.receiveCANbyID(0x580 + self.CANAddr_AO,
                                                                             self.waiting_time,1)
                    if bool_receive:
                        # # print(f'm_can_obj.Data[4]:{self.m_can_obj.Data[4]}')
                        # # print(f'm_can_obj.Data[5]:{self.m_can_obj.Data[5]}')
                        # # print(f'AORangeArray[0]:{self.AORangeArray[0]}')
                        # # print(f'AORangeArray[1]:{self.AORangeArray[1]}')
                        if hex(self.m_can_obj.Data[4]) == hex(self.cur_AORangeArray[2*typeNum]) and hex(
                                self.m_can_obj.Data[5]) == hex(self.cur_AORangeArray[2*typeNum+1]):
                            # print(f'{AOChannel}.成功设置AO通道{AOChannel}的量程为”{self.cur_name_array[typeNum]}“。\n\n')
                            self.pauseOption()
                            if not self.is_running:
                                return False
                            self.result_signal.emit(f'{AOChannel}.成功设置AO通道{AOChannel}的量程为”'
                                                    f'{self.cur_name_array[typeNum]}“。\n\n')
                            break
                        else:
                            self.pauseOption()
                            if not self.is_running:
                                return False
                            self.result_signal.emit(f'{AOChannel}.未成功设置AO通道{AOChannel}的量程为”'
                                                    f'{self.cur_name_array[typeNum]}“。\n\n')
                            # print(f'{AOChannel}.未成功设置AO通道{AOChannel}的量程为”'
                            #       f'{self.cur_name_array[typeNum]}“。\n\n')
                            return False
                else:
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    self.result_signal.emit(f'{AOChannel}.未成功设置AO通道{AOChannel}的量程为”'
                                            f'{self.cur_AIRangeArray[typeNum]}“。\n\n')
                    return False

        if AOChannel == 4:
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(self.HORIZONTAL_LINE)
        return True

    # 检查配套AI模块输入类型是否正确
    def setAIInputType(self, AIChannel, type:str, typeNum:int):
        # CAN_option.close(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX)
        # self.can_start()
        if AIChannel == 1:
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(f'设置AI通道量程：\n\n')

        self.m_transmitData[0] = 0x2b
        self.m_transmitData[1] = 0x10
        self.m_transmitData[2] = 0x61
        self.m_transmitData[3] = AIChannel
        self.m_transmitData[6] = 0x00
        self.m_transmitData[7] = 0x00

        self.clearList(self.m_can_obj.Data)

        if type == 'AOVoltage' or type == 'AIVoltage':
            self.m_transmitData[4] = self.vol_AIRangeArray[2*typeNum]
            self.m_transmitData[5] = self.vol_AIRangeArray[2*typeNum + 1]
            while True:
                # QApplication.processEvents()
                if CAN_option.transmitCAN(0x600 + self.CANAddr_AI, self.m_transmitData,1)[0]:
                    bool_receive, self.m_can_obj = CAN_option.receiveCANbyID(0x580 + self.CANAddr_AI,
                                                                             self.waiting_time,1)
                    if bool_receive:
                        # # print(f'm_can_obj.Data[4]:{self.m_can_obj.Data[4]}')
                        # # print(f'm_can_obj.Data[5]:{self.m_can_obj.Data[5]}')
                        # # print(f'self.vol_AIRangeArray[2*typeNum]:{self.vol_AIRangeArray[2*typeNum]')
                        # # print(f'self.vol_AIRangeArray[2*typeNum]+1:self.vol_AIRangeArray[2*typeNum+1]')
                        if hex(self.m_can_obj.Data[4]) == hex(self.vol_AIRangeArray[2*typeNum]) and hex(
                                self.m_can_obj.Data[5]) == hex(self.vol_AIRangeArray[2*typeNum+1]):
                            # print(f'{AIChannel}.成功设置AI通道{AIChannel}的量程为"{self.vol_name_array[typeNum]}"。\n\n')

                            self.pauseOption()
                            if not self.is_running:
                                return False
                            self.result_signal.emit(f'{AIChannel}.成功设置AI通道{AIChannel}的量程为"'
                                                    f'{self.vol_name_array[typeNum]}"。\n\n')
                            break
                        else:
                            self.pauseOption()
                            if not self.is_running:
                                return False
                            self.result_signal.emit(f'{AIChannel}.未成功设置AI通道{AIChannel}的量程为"'
                                                    f'{self.vol_name_array[typeNum]}"。\n\n')
                            return False
                else:
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    self.result_signal.emit(f'{AIChannel}.未成功设置AI通道{AIChannel}的量程为”'
                                            f'{self.vol_name_array[typeNum]}“。\n\n')
                    return False
                    # # print(f'{AIChannel}.未成功设置AI通道{AIChannel}的量程为”电压+/-10V“。\n\n')
        elif type == 'AOCurrent' or type == 'AICurrent':
            self.m_transmitData[4] = self.cur_AIRangeArray[2*typeNum]
            self.m_transmitData[5] = self.cur_AIRangeArray[2*typeNum + 1]
            while True:
                # self.isPause()
                # if not self.isStop():
                #                  return
                QApplication.processEvents()
                if CAN_option.transmitCAN(0x600 + self.CANAddr_AI, self.m_transmitData,1)[0]:
                    bool_receive, self.m_can_obj = CAN_option.receiveCANbyID(0x580 + self.CANAddr_AI,
                                                                             self.waiting_time,1)
                    if bool_receive:
                        # # print(f'm_can_obj.Data[4]:{self.m_can_obj.Data[4]}')
                        # # print(f'm_can_obj.Data[5]:{self.m_can_obj.Data[5]}')
                        # # print(f'AIRangeArray[0]:{self.AIRangeArray[0]}')
                        # # print(f'AIRangeArray[1]:{self.AIRangeArray[1]}')
                        if hex(self.m_can_obj.Data[4]) == hex(self.cur_AIRangeArray[2*typeNum]) and hex(
                                self.m_can_obj.Data[5]) == hex(self.cur_AIRangeArray[2*typeNum+1]):
                            # print(f'{AIChannel}.成功设置AI通道{AIChannel}的量程为”{self.cur_name_array[typeNum]}“。\n\n')

                            self.pauseOption()
                            if not self.is_running:
                                return False
                            self.result_signal.emit(f'{AIChannel}.成功设置AI通道{AIChannel}的量程为”'
                                                    f'{self.cur_name_array[typeNum]}“。\n\n')
                            break
                        else:
                            self.pauseOption()
                            if not self.is_running:
                                return False
                            self.result_signal.emit(f'{AIChannel}.未成功设置AI通道{AIChannel}的量程为'
                                                    f'”{self.cur_name_array[typeNum]}“。\n\n')
                            # print(f'{AIChannel}.未成功设置AI通道{AIChannel}的量程为”{self.cur_name_array[typeNum]}“。\n\n')
                            return False
                else:
                    self.pauseOption()
                    if not self.is_running:
                        return False
                    self.result_signal.emit(f'{AIChannel}.未成功设置AI通道{AIChannel}的量程为”'
                                            f'{self.cur_name_array[typeNum]}“。\n\n')
                    return False
        if AIChannel == 4:
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit(self.HORIZONTAL_LINE)

        return True

    def setAIChOutCalibrate(self):
        bool_all = True
        self.m_transmitData = [0x23, 0x3f, 0x64, 0x00, 0x45, 0x58, 0x49, 0x54]

        bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + self.CANAddr_AI), self.m_transmitData,1)
        bool_all = bool_all & bool_transmit

        return bool_all

    def normal_writeValuetoAO(self,value):
        bool_all = True
        for i in range(self.m_Channels):
            self.m_transmitData = [0x2b,0x11,0x64,i+1,(value & 0xff),((value >> 8) & 0xff),0x00,0x00]

            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600+self.CANAddr_AO), self.m_transmitData,1)
            bool_all = bool_all & bool_transmit

        return bool_all

    def channelZero(self):
        self.pauseOption()
        if not self.is_running:
            return False
        self.result_signal.emit(f'对所有通道值进行归零处理' + self.HORIZONTAL_LINE)
        # # print('1111111111111111')
        isZero = self.normal_writeValuetoAO(0)
        # # print('2222222222222222')
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

    def testRunErr(self,addr):
        self.testNum -= 1
        self.isLEDRunOK = True
        self.isLEDErrOK = True
        self.isLEDPass = True

        self.pauseOption()
        if not self.is_running:
            return False
        self.item_signal.emit([1,1,0,''])
            # mTable.item(1, i).setBackground(QtGui.QColor(255, 255, 0))
            # if i == 1:
            #     mTable.item(1, i).setText('正在检测')
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
        bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr),self.m_transmitData,1)
        runEnd_time = time.time()
        runTest_time = round(runEnd_time - runStart_time,2)
        time.sleep(0.5)
        # reply = QMessageBox.question(None, '检测RUN &ERROR', 'RUN指示灯是否点亮？',
        #                              QMessageBox.AcceptRole | QMessageBox.RejectRole,
        #                              QMessageBox.AcceptRole)
        image_RUN = self.current_dir + '/AE_RUN.png'
        self.pic_messageBox_signal.emit(['检测RUN &ERROR', 'RUN指示灯是否如图所示点亮(绿灯)？', image_RUN])
        reply = self.result_queue.get()
        if reply == QMessageBox.AcceptRole:
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
            self.item_signal.emit([1,2,1,f'{runTest_time}'])
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("LED RUN 测试通过\n关闭LED RUN\n"+self.HORIZONTAL_LINE)
            # print("LED RUN 测试通过\n关闭LED RUN\n"+self.HORIZONTAL_LINE)
            self.clearList(self.m_transmitData)
            self.m_transmitData[0] = 0x2f
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x01
            self.m_transmitData[4] = 0x00
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData,1)
            time.sleep(1)
            self.isLEDRunOK = True
            # self.result_signal.emit(f'self.isLEDRunOK:{self.isLEDRunOK}')
        elif reply == QMessageBox.RejectRole:
            self.runLED = False
            self.pauseOption()
            if not self.is_running:
                return False
            self.item_signal.emit([1,2,2,f'{runTest_time}'])
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

        bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData,1)
        errorEnd_time = time.time()
        errorTest_time = round(errorEnd_time-errorStart_time,2)
        time.sleep(0.5)
        image_ERR = self.current_dir + '/AE_ERROR.png'
        self.pic_messageBox_signal.emit(['检测RUN &ERROR', 'ERROR指示灯是否如图所示点亮（红灯）？', image_ERR])
        reply = self.result_queue.get()
        # reply = QMessageBox.question(None, '检测RUN &ERROR', 'ERROR指示灯是否点亮？',
        #                              QMessageBox.AcceptRole | QMessageBox.RejectRole,
        #                              QMessageBox.AcceptRole)
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
            self.item_signal.emit([2,2,1,f'{errorTest_time}'])
            self.pauseOption()
            if not self.is_running:
                return False
            self.result_signal.emit("LED ERROR 测试通过\n关闭LED ERROR\n"+self.HORIZONTAL_LINE)
            # print("LED ERROR 测试通过\n关闭LED ERROR\n"+self.HORIZONTAL_LINE)
            self.clearList(self.m_transmitData)
            self.m_transmitData[0] = 0x2f
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x02
            self.m_transmitData[4] = 0x00
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData,1)
            time.sleep(1)
            self.isLEDErrOK = True
        elif reply == QMessageBox.RejectRole:
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
            # print("LED ERROR 测试不通过\n" + self.HORIZONTAL_LINE)
            self.isLEDErrOK = False

        self.isLEDPass = self.isLEDPass & self.isLEDErrOK

        return True

    def testCANRunErr(self, addr):
        self.testNum -= 1
        # 关闭CAN设备
        #CAN_option.close(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX)
        # 启动CAN设备并打开CAN通道
        #self.can_start()
        self.isLEDCANRunOK = True
        self.isLEDCANErrOK = True
        self.isLEDCANPass = True
        # reply = QMessageBox.question(None, '检测CAN_RUN &CAN_ERROR', '是否开始进行CAN_RUN 和CAN_ERROR 检测？',
        #                              QMessageBox.AcceptRole | QMessageBox.RejectRole,
        #                              QMessageBox.AcceptRole)
        # if self.tabIndex == 0:
        #     mTable = self.tableWidget_DIDO
        # elif self.tabIndex == 1:
        #     mTable = self.tableWidget_AI
        # elif self.tabIndex == 2:
        #     mTable = self.tableWidget_AO
        # if reply == QMessageBox.AcceptRole:
        CANRunStart_time = time.time()
        self.pauseOption()
        if not self.is_running:
            return False
        self.item_signal.emit([3,1,0,''])

        self.pauseOption()
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
        bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData,1)
        CANRunEnd_time = time.time()
        CANRunTest_time = round(CANRunEnd_time - CANRunStart_time,2)
        time.sleep(0.5)
        self.messageBox_signal.emit(['检测CAN_RUN &CAN_ERROR', 'CAN_RUN指示灯是否点亮（绿灯）？'])
        reply = self.result_queue.get()
        # reply = QMessageBox.question(None, '检测CAN_RUN &CAN_ERROR', '指示灯是否点亮？',
        #                              QMessageBox.AcceptRole | QMessageBox.RejectRole,
        #                              QMessageBox.AcceptRole)
        if reply == QMessageBox.AcceptRole:
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
            # print("LED CAN_RUN 测试通过\n关闭LED CAN_RUN\n" + self.HORIZONTAL_LINE)
            self.clearList(self.m_transmitData)
            self.m_transmitData[0] = 0x2f
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x03
            self.m_transmitData[4] = 0x00
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData,1)
            time.sleep(1)
            self.isLEDCANRunOK = True
        elif reply == QMessageBox.RejectRole:
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
            # print("LED CAN_RUN 测试不通过\n" + self.HORIZONTAL_LINE)
            self.isLEDCANRunOK = False
        self.isLEDCANPass = self.isLEDCANPass & self.isLEDCANRunOK

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
        # print("2.进行 LED CAN_ERROR 测试\n\n")
        self.clearList(self.m_transmitData)
        self.m_transmitData[0] = 0x2f
        self.m_transmitData[1] = 0xf6
        self.m_transmitData[2] = 0x5f
        self.m_transmitData[3] = 0x04
        self.m_transmitData[4] = 0x01
        bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData,1)
        CANErrEnd_time = time.time()
        CANErrTest_time = round(CANErrEnd_time - CANErrStart_time,2)
        time.sleep(0.5)
        self.messageBox_signal.emit(['检测CAN_RUN &CAN_ERROR', 'CAN_ERROR指示灯是否点亮（红灯）？'])
        reply = self.result_queue.get()
        # reply = QMessageBox.question(None, '检测CAN_RUN &CAN_ERROR', '指示灯是否点亮？',
        #                              QMessageBox.AcceptRole | QMessageBox.RejectRole,
        #                              QMessageBox.AcceptRole)
        if reply == QMessageBox.AcceptRole:
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
            # print("LED CAN_ERROR 测试通过\n关闭LED CAN_ERROR\n" + self.HORIZONTAL_LINE)
            self.clearList(self.m_transmitData)
            self.m_transmitData[0] = 0x2f
            self.m_transmitData[1] = 0xf6
            self.m_transmitData[2] = 0x5f
            self.m_transmitData[3] = 0x04
            self.m_transmitData[4] = 0x00
            bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + addr), self.m_transmitData,1)
            time.sleep(1)
            self.isLEDCANErrOK = True
        elif reply == QMessageBox.RejectRole:
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
            # print("LED CAN_ERROR 测试不通过\n" + self.HORIZONTAL_LINE)
            self.isLEDCANErrOK = False
        self.isLEDCANPass = self.isLEDCANPass & self.isLEDCANErrOK

        return True


    def fillInAIData(self,station, book, sheet):
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
            sheet.write(self.generalTest_row, 6, '未测试', warning_style)
            sheet.write(self.generalTest_row, 9, '未测试', warning_style)
        if not self.isTestCANRunErr:
            sheet.write(self.generalTest_row, 12, '未测试', warning_style)
            sheet.write(self.generalTest_row, 15, '未测试', warning_style)
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

        self.vol_excelName_array = ["-10V～10V", "-5V～5V", "0V～5V", "0V～10V", "1V～5V"]
        self.cur_excelName_array = ["4mA～20mA", "0mA～20mA"]
        # 填写信号类型、通道号、测试点数据
        if self.isAITestVol:
            all_row = 9 + 4 + 4 + (2 + self.AI_Channels*5) + 2  # CPU + DI + DO + AI + AO
            sheet.write_merge(self.AI_row + 2, self.AI_row + 1 + self.AI_Channels*5, 1, 1, '电压', pass_style)
            for typeNum in range(5):
                sheet.write_merge(self.AI_row + 2 + self.AI_Channels*typeNum,
                                  self.AI_row + 1 + self.AI_Channels*(typeNum+1),2, 2,
                                  f'{self.vol_excelName_array[typeNum]}', pass_style)
                for i in range(5):
                    for j in range(self.AI_Channels):
                        # 通道号
                        sheet.write(self.AI_row + 2 + j + self.AI_Channels * typeNum, 3, f'CH{j + 1}', pass_style)
                        # 理论值
                        sheet.write(self.AI_row + 2 + j + self.AI_Channels * typeNum, 3 + 3 * i + 1,
                                    f'{int(self.volValue_array[typeNum][i])}', pass_style)
                        # # 测试值
                        # sheet.write(self.AI_row + 2 + j, 4 + 3 * i + 1, f'{self.volReceValue[typeNum][i][j]}', pass_style)
                        # 精度
                        if isinstance(self.volPrecision[typeNum][i][j], float) and \
                                abs(self.volPrecision[typeNum][i][j]) < 1:
                            # 测试值
                            sheet.write(self.AI_row + 2 + j + self.AI_Channels * typeNum, 4 + 3 * i + 1,
                                        f'{self.volReceValue[typeNum][i][j]}', pass_style)
                            sheet.write(self.AI_row + 2 + j + self.AI_Channels * typeNum, 5 + 3 * i + 1,
                                        f'{self.volPrecision[typeNum][i][j]}‰', pass_style)
                        elif isinstance(self.volPrecision[typeNum][i][j], str) and \
                                self.volPrecision[typeNum][i][j] == '-':
                            # 测试值
                            sheet.write(self.AI_row + 2 + j + self.AI_Channels * typeNum, 4 + 3 * i + 1,
                                        f'{self.volReceValue[typeNum][i][j]}', warning_style)
                            sheet.write(self.AI_row + 2 + j + self.AI_Channels * typeNum, 5 + 3 * i + 1,
                                        f'{self.volPrecision[typeNum][i][j]}', warning_style)
                            self.errorNum += 1
                            self.errorInf += f'\n{self.errorNum})AI模块量程"{self.vol_excelName_array[typeNum]}"' \
                                             f'的测试点{i + 1}在通道{j + 1}的数据接收有误'
                        elif isinstance(self.volPrecision[typeNum][i][j], float) and \
                                abs(self.volPrecision[typeNum][i][j]) >= 1:
                            self.errorNum += 1
                            self.errorInf += f'\n{self.errorNum})AI模块量程"{self.vol_excelName_array[typeNum]}"' \
                                             f'的测试点{i + 1}在通道{j + 1}的电压精度超出范围'
                            # 测试值
                            sheet.write(self.AI_row + 2 + j + self.AI_Channels * typeNum, 4 + 3 * i + 1,
                                        f'{self.volReceValue[typeNum][i][j]}', fail_style)
                            sheet.write(self.AI_row + 2 + j+ self.AI_Channels * typeNum, 5 + 3 * i + 1,
                                        f'{self.volPrecision[typeNum][i][j]}‰', fail_style)
        if self.isAITestVol and self.isAITestCur:
            all_row = 9 + 4 + 4 + (2 + 7 * self.AI_Channels) + 2  # CPU + DI + DO + AI + AO
            sheet.write_merge(self.AI_row + 2 + self.AI_Channels*5, self.AI_row + 1 + 7 * self.AI_Channels, 1, 1,
                              '电流', pass_style)
            # sheet.write_merge(self.AI_row + 2, self.AI_row + 1 + self.AI_Channels * 5, 1, 1, '电压', pass_style)
            for typeNum in range(2):
                sheet.write_merge(self.AI_row + 2 + self.AI_Channels * 5 + self.AI_Channels*typeNum,
                                  self.AI_row + 1 + self.AI_Channels * 5 + self.AI_Channels*(typeNum+1),2, 2,
                                  f'{self.cur_excelName_array[typeNum]}', pass_style)
                for i in range(5):
                    for j in range(self.AI_Channels):
                        # 通道号
                        sheet.write(self.AI_row + 2 + 5*self.AI_Channels + j + self.AI_Channels*typeNum,
                                    3,f'CH{j + 1}',pass_style)
                        # 理论值
                        sheet.write(self.AI_row + 2 + 5*self.AI_Channels + j + self.AI_Channels*typeNum,
                                    3 + 3 * i + 1,f'{int(self.curValue_array[typeNum][i])}',pass_style)
                        # # 测试值
                        # sheet.write(self.AI_row + 2 + 5*self.AI_Channels + j + self.AI_Channels*typeNum,
                        #             4 + 3 * i + 1,f'{self.curReceValue[typeNum][i][j]}',pass_style)
                        # 精度
                        if isinstance(self.curPrecision[typeNum][i][j], float) and \
                                abs(self.curPrecision[typeNum][i][j]) < 1:
                            # 测试值
                            sheet.write(self.AI_row + 2 + 5 * self.AI_Channels + j + self.AI_Channels * typeNum,
                                        4 + 3 * i + 1, f'{self.curReceValue[typeNum][i][j]}', pass_style)
                            # 精度
                            sheet.write(self.AI_row + 2 + 5*self.AI_Channels + j + self.AI_Channels*typeNum,
                                        5 + 3 * i + 1,f'{self.curPrecision[typeNum][i][j]}‰', pass_style)
                        elif isinstance(self.curPrecision[typeNum][i][j], str) and \
                                self.curPrecision[typeNum][i][j] == '-':
                            # 测试值
                            sheet.write(self.AI_row + 2 + 5 * self.AI_Channels + j + self.AI_Channels * typeNum,
                                        4 + 3 * i + 1, f'{self.curReceValue[typeNum][i][j]}', warning_style)
                            # 精度
                            sheet.write(self.AI_row + 2 + 5 * self.AI_Channels + j + self.AI_Channels * typeNum,
                                        5 + 3 * i + 1,f'{self.curPrecision[typeNum][i][j]}', warning_style)
                            self.errorNum += 1
                            self.errorInf += f'\n{self.errorNum})AI模块量程"{self.cur_excelName_array[typeNum]}"' \
                                             f'的测试点{i + 1}在通道{j + 1}的数据接收有误'
                        elif isinstance(self.curPrecision[typeNum][i][j], float) and \
                                abs(self.curPrecision[typeNum][i][j]) >= 1:
                            self.errorNum += 1
                            self.errorInf += f'\n{self.errorNum})AI模块量程"{self.cur_excelName_array[typeNum]}"' \
                                             f'的测试点{i + 1}在通道{j + 1}的电流精度超出范围'
                            # 测试值
                            sheet.write(self.AI_row + 2 + 5 * self.AI_Channels + j + self.AI_Channels * typeNum,
                                        4 + 3 * i + 1, f'{self.curReceValue[typeNum][i][j]}', fail_style)
                            sheet.write(self.AI_row + 2 + 5*self.AI_Channels + j + self.AI_Channels*typeNum,
                                        5 + 3 * i + 1,f'{self.curPrecision[typeNum][i][j]}‰', fail_style)
        if not self.isAITestVol and self.isAITestCur:
            all_row = 9 + 4 + 4 + (2 + 2*self.AI_Channels) + 2  # CPU + DI + DO + AI + AO
            sheet.write_merge(self.AI_row + 2, self.AI_row + 1 + 2 * self.AI_Channels, 1, 1, '电流', pass_style)
            for typeNum in range(2):
                sheet.write_merge(self.AI_row + 2 + self.AI_Channels * typeNum,
                                  self.AI_row + 1 + self.AI_Channels * (typeNum + 1), 2, 2,
                                  f'{self.cur_excelName_array[typeNum]}', pass_style)
                for i in range(5):
                    for j in range(self.AI_Channels):
                        # 通道号
                        sheet.write(self.AI_row + 2 + j + self.AI_Channels * typeNum, 3, f'CH{j + 1}', pass_style)
                        # 理论值
                        sheet.write(self.AI_row + 2 + j + self.AI_Channels * typeNum, 3 + 3 * i + 1,
                                    f'{int(self.curValue_array[typeNum][i])}', pass_style)
                        # # 测试值
                        # sheet.write(self.AI_row + 2 + j + self.AI_Channels*typeNum, 4 + 3 * i + 1,
                        #             f'{self.curReceValue[typeNum][i][j]}', pass_style)
                        # 精度
                        if isinstance(self.curPrecision[typeNum][i][j], float) and \
                                abs(self.curPrecision[typeNum][i][j]) < 1:
                            # 测试值
                            sheet.write(self.AI_row + 2 + j + self.AI_Channels * typeNum, 4 + 3 * i + 1,
                                        f'{self.curReceValue[typeNum][i][j]}', pass_style)
                            # 精度
                            sheet.write(self.AI_row + 2 + j + self.AI_Channels * typeNum, 5 + 3 * i + 1,
                                        f'{self.curPrecision[typeNum][i][j]}‰', pass_style)
                        elif isinstance(self.curPrecision[typeNum][i][j], str) and \
                                self.curPrecision[typeNum][i][j] == '-':
                            # 测试值
                            sheet.write(self.AI_row + 2 + j + self.AI_Channels * typeNum, 4 + 3 * i + 1,
                                        f'{self.curReceValue[typeNum][i][j]}', warning_style)
                            # 精度
                            sheet.write(self.AI_row + 2 + j + self.AI_Channels * typeNum, 5 + 3 * i + 1,
                                        f'{self.volPrecision[typeNum][i][j]}', warning_style)
                            self.errorNum += 1
                            self.errorInf += f'\n{self.errorNum})AI模块量程"{self.cur_excelName_array[typeNum]}"' \
                                             f'的测试点{i + 1}在通道{j + 1}的数据接收有误'
                        elif isinstance(self.curPrecision[typeNum][i][j], float) and \
                                abs(self.curPrecision[typeNum][i][j]) >= 1:
                            self.errorNum += 1
                            self.errorInf += f'\n{self.errorNum})AI模块量程"{self.cur_excelName_array[typeNum]}"' \
                                             f'的测试点{i + 1}在通道{j + 1}的电流精度超出范围'
                            # 测试值
                            sheet.write(self.AI_row + 2 + j + self.AI_Channels * typeNum, 4 + 3 * i + 1,
                                        f'{self.curReceValue[typeNum][i][j]}', fail_style)
                            sheet.write(self.AI_row + 2 + j + self.AI_Channels * typeNum, 5 + 3 * i + 1,
                                        f'{self.curPrecision[typeNum][i][j]}‰', fail_style)

        elif not self.isAITestVol and not self.isAITestCur:
            all_row = 9 + 4 + 4 + 2 + 2  # CPU + DI + DO + AI + AO
        # # print(f'self.isAIPassTest:{self.isAIPassTest}')
        self.isAIPassTest = (((((self.isAIPassTest & self.isLEDRunOK) & self.isLEDErrOK)
                               & self.CAN_runLED) & self.CAN_errorLED) & self.appearance)
        # self.result_signal.emit(f'self.isLEDRunOK:{self.isLEDRunOK}')
        # # print(f'self.isAIPassTest:{self.isAIPassTest}')
        # # print(f'self.isLEDRunOK:{self.isLEDRunOK}')
        # # print(f'self.isLEDErrOK:{self.isLEDErrOK}')
        # # print(f'self.CAN_runLED:{self.CAN_runLED}')
        # # print(f'self.CAN_errorLED:{self.CAN_errorLED}')
        # # print(f'self.appearance:{self.appearance}')
        # # print(f'self.testNum:{self.testNum}')
        # self.result_signal.emit(f'self.testNum:{self.testNum}')
        name_save = ''
        if self.isAIPassTest and self.testNum == 0:
            name_save = '合格'
            sheet.write(self.generalTest_row + all_row + 1, 4, '■ 合格', pass_style)
            sheet.write(self.generalTest_row + all_row + 1, 6,
                        '------------------ 全部项目测试通过！！！ ------------------', pass_style)
            self.label_signal.emit(['pass','全部通过'])
            self.print_signal.emit([f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}', 'PASS', ''])
            # self.label.setStyleSheet(self.testState_qss['pass'])
            # self.label.setText('全部通过')
        elif self.isAIPassTest and self.testNum > 0:
            name_save = '部分合格'
            sheet.write(self.generalTest_row + all_row + 1, 4, '■ 部分合格', pass_style)
            sheet.write(self.generalTest_row + all_row + 1, 6,
                        '------------------ 注意：有部分项目未测试！！！ ------------------', warning_style)
            self.label_signal.emit(['testing', '部分通过'])
            self.print_signal.emit([f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}', 'PASS', ''])
            # self.label.setStyleSheet(self.testState_qss['testing'])
            # self.label.setText('部分通过')
        elif not self.isAIPassTest:
            name_save = '不合格'
            sheet.write(self.generalTest_row + all_row + 2, 4, '■ 不合格', fail_style)
            sheet.write(self.generalTest_row + all_row + 2, 6, f'不合格原因：{self.errorInf}', fail_style)
            self.label_signal.emit(['fail', '未通过'])
            self.print_signal.emit(
                [f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}', 'FAIL', self.errorInf])
            # self.label.setStyleSheet(self.testState_qss['fail'])
            # self.label.setText('未通过')

        self.saveExcel_signal.emit([book,f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}.xls'])
        # book.save(self.saveDir + f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}.xls')

    def initPara_array(self):
        # 5个量程 -> 每个量程5个点 -> 每个点4个通道
        self.volReceValue = [[['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'],
                         ['-', '-', '-', '-']],
                        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'],
                         ['-', '-', '-', '-']],
                        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'],
                         ['-', '-', '-', '-']],
                        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'],
                         ['-', '-', '-', '-']],
                        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'],
                         ['-', '-', '-', '-']]]

        self.volPrecision = [[['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'],
                         ['-', '-', '-', '-']],
                        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'],
                         ['-', '-', '-', '-']],
                        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'],
                         ['-', '-', '-', '-']],
                        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'],
                         ['-', '-', '-', '-']],
                        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'],
                         ['-', '-', '-', '-']]]
        # 2个量程 -> 每个量程5个点 -> 每个点4个通道
        self.curReceValue = [[['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'],
                         ['-', '-', '-', '-']],
                        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'],
                         ['-', '-', '-', '-']]]

        self.curPrecision = [[['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'],
                         ['-', '-', '-', '-']],
                        [['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'], ['-', '-', '-', '-'],
                         ['-', '-', '-', '-']]]
    def pause_work(self):
        self.is_pause = True

    def resume_work(self):
        self.is_pause = False
        
    def pauseOption(self):
        if self.is_pause:
            while True:
                if self.pause_num == 1 and self.is_running:
                    self.pause_num += 1
                    # self.result_signal.emit(self.HORIZONTAL_LINE + '暂停中…………' + self.HORIZONTAL_LINE)
                if not self.is_pause:
                    self.pause_num = 1
                    break
                
    def stop_work(self):
        self.resume_work()
        self.is_running = False
        self.setAIChOutCalibrate()
        self.label_signal.emit(['fail', '测试停止'])

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
    #     if (self.isAITestVol and not self.isAITestCur):
    #         sheet.write_merge(self.AI_row, self.AI_row + 1 + self.AI_Channels*5, 0, 0, 'AI信号', leftTitle_style)
    #         self.AO_row = self.AI_row + 2 + self.AI_Channels*5
    #     elif  not self.isAITestVol and self.isAITestCur:
    #         sheet.write_merge(self.AI_row, self.AI_row + 1 + self.AI_Channels * 2, 0, 0, 'AO信号', leftTitle_style)
    #         self.AO_row = self.AI_row + 2 + self.AI_Channels * 2
    #     elif self.isAITestVol and self.isAITestCur:
    #         sheet.write_merge(self.AI_row, self.AI_row + 1 + 7 * self.AI_Channels, 0, 0, 'AI信号', leftTitle_style)
    #         self.AO_row = self.AI_row + 2 + 7 * self.AI_Channels
    #     elif not self.isAITestVol and not self.isAITestCur:
    #         sheet.write_merge(self.AI_row, self.AI_row + 1, 0, 0, 'AI信号', leftTitle_style)
    #         self.AO_row = self.AI_row + 2
    #
    #     # sheet.write_merge(self.AO_row, self.AO_row + 1 + self.AO_Channels, 0, 0, 'AO信号', leftTitle_style)
    #     sheet.write_merge(self.AO_row, self.AO_row + 1, 0, 0, 'AO信号', leftTitle_style)
    #     self.result_row = self.AO_row + 2
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
    #     sheet.write_merge(self.AI_row, self.AI_row + 1, 2, 2, '量程', contentTitle_style)
    #     sheet.write_merge(self.AI_row, self.AI_row + 1, 3, 3, '通道号', contentTitle_style)
    #     sheet.write_merge(self.AI_row, self.AI_row, 3 + 1, 5 + 1, '测试点1(100%)', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 3 + 1, '理论值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 4 + 1, '测试值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 5 + 1, '精度', contentTitle_style)
    #
    #     sheet.write_merge(self.AI_row, self.AI_row, 6 + 1, 8 + 1, '测试点2(75%)', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 6 + 1, '理论值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 7 + 1, '测试值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 8 + 1, '精度', contentTitle_style)
    #
    #     sheet.write_merge(self.AI_row, self.AI_row, 9 + 1, 11 + 1, '测试点3(50%)', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 9 + 1, '理论值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 10 + 1, '测试值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 11 + 1, '精度', contentTitle_style)
    #
    #     sheet.write_merge(self.AI_row, self.AI_row, 12 + 1, 14 + 1, '测试点4(25%)', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 12 + 1, '理论值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 13 + 1, '测试值', contentTitle_style)
    #     sheet.write(self.AI_row + 1, 14 + 1, '精度', contentTitle_style)
    #
    #     sheet.write_merge(self.AI_row, self.AI_row, 15 + 1, 17 + 1, '测试点5(0%)', contentTitle_style)
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
    #     self.fillInAIData(station, book, sheet)
    #     # if module == 'DI':
    #     #     self.fillInDIData(station, book, sheet)
    #     # elif module == 'DO':
    #     #     self.fillInDOData(station, book, sheet)
    #     # elif module == 'AI':
    #     #     # # print('打印AI检测结果')
    #     #     self.fillInAIData(station, book, sheet)
    #     # elif module == 'AO':
    #     #     # # print('打印AI检测结果')
    #     #     self.fillInAOData(station, book, sheet)
