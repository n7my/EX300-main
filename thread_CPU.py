# -*- coding: utf-8 -*-

import time

from PyQt5.QtCore import QThread, pyqtSignal
from main_logic import *
from CAN_option import *
import otherOption
import logging

class CPUThread(QObject):
    result_signal = pyqtSignal(str)
    item_signal = pyqtSignal(list)
    pass_signal = pyqtSignal(bool)

    messageBox_signal = pyqtSignal(list)
    # excel_signal = pyqtSignal(list)
    allFinished_signal = pyqtSignal()
    label_signal = pyqtSignal(list)
    saveExcel_signal = pyqtSignal(list)
    print_signal = pyqtSignal(list)

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
    serial_transmitData =[]
    serial_receiveData = []
    # 主副线程弹窗结果
    result_queue = 0

    #各测试项通过标志
    isPassAppearance = False
    isPassTypeTest = False
    isPassSRAM = False
    isPassFLASH = False
    isPassConfig = False
    isPassFPGA = False
    isPassLever = False
    isPassMFK = False
    isPassRTC = False
    isPassPowerOffSave = False
    isPassLED = False
    isPassIn = False
    isPassOut =False
    isPassETH  = False
    isPass232 = False
    isPass485 = False
    isPassRightCAN = False
    isPassEX = False

    #是否取消后续所有测试
    isCancelAllTest = False



    def __init__(self, inf_CPUlist: list, result_queue):
        super().__init__()
        self.result_queue = result_queue
        self.is_running = True
        self.is_pause = False
        logging.basicConfig(filename='example.log', level=logging.DEBUG)
        # self.inf_CPUlist = [self.inf_param,self.inf_product, self.inf_CANIPAdrr,
        #                                 self.inf_serialPort, self.inf_CPU_test]
        # self.inf_param
        self.mTable = inf_CPUlist[0][0]
        self.module_1 = inf_CPUlist[0][1]
        self.module_2 = inf_CPUlist[0][2]
        self.module_3 = inf_CPUlist[0][3]
        self.module_4 = inf_CPUlist[0][4]
        self.module_5 = inf_CPUlist[0][5]
        self.testNum = inf_CPUlist[0][6]
        # self.inf_product = [self.module_type, self.module_pn, self.module_sn,
        #                    self.module_rev, self.in_Channels,self.out_Channels]
        #获取产品信息
        self.module_type = inf_CPUlist[1][0]
        self.module_pn = inf_CPUlist[1][1]
        self.module_sn = inf_CPUlist[1][2]
        self.module_rev = inf_CPUlist[1][3]
        self.module_MAC = inf_CPUlist[1][4]
        self.in_Channels = int(inf_CPUlist[1][5])
        self.out_Channels = int(inf_CPUlist[1][6])
        # self.inf_CANIPAdrr = [self.CANAddr1,self.CANAddr2,self.CANAddr3,self.CANAddr4,
        #                                   self.CANAddr5,self.IPAddr]
        # 获取CAN、IP地址
        self.CANAddr1 = inf_CPUlist[2][0]
        self.CANAddr2 = inf_CPUlist[2][1]
        self.CANAddr3 = inf_CPUlist[2][2]
        self.CANAddr4 = inf_CPUlist[2][3]
        self.CANAddr5 = inf_CPUlist[2][4]
        # self.IPAddr = inf_CPUlist[2][5]
        # self.inf_serialPort = [self.serialPort_232, self.serialPort_485,
        #                       self.serialPort_typeC, self.saveDir]
        #获取串口信息
        self.serialPort_232 = inf_CPUlist[3][0]
        self.serialPort_485 = inf_CPUlist[3][1]
        self.serialPort_typeC = inf_CPUlist[3][2]
        self.saveDir = inf_CPUlist[3][3]

        # 获取CPU检测信息
        self.CPU_isTest = inf_CPUlist[4]

        #信号等待时间(ms)
        self.waiting_time = 3000

        self.isCancelAllTest = False
    def CPUOption(self):
        self.isExcel = True
        #测试是否成功标志
        # self.testSign = True
        try:
            ############################################################总线初始化############################################################
            try:
                time_online = time.time()
                while True:
                    QApplication.processEvents()
                    if (time.time() - time_online) * 1000 > self.waiting_time:
                        self.pauseOption()
                        if not self.is_running:
                            # 后续测试全部取消
                            self.cancelAllTest()
                            # self.isTest = False
                            # self.CPU_isTest = [False for x in range(self.testNum)]
                            # self.testSign = False
                            break
                        self.result_signal.emit(f'错误：总线初始化超时！' + self.HORIZONTAL_LINE)
                        self.messageBox_signal.emit(['错误提示', '总线初始化超时！请检查CAN分析仪或各设备是否正确连接！'])
                        # 后续测试全部取消
                        self.cancelAllTest()
                        # self.isTest = False
                        # self.CPU_isTest = [False for x in range(self.testNum)]
                        # self.testSign = False
                        break
                    CANAddr_array = [self.CANAddr1,self.CANAddr2,self.CANAddr3,self.CANAddr4,self.CANAddr5]
                    module_array = [self.module_1, self.module_2, self.module_3, self.module_4, self.module_5]
                    bool_online, eID = otherOption.isModulesOnline(CANAddr_array,module_array, self.waiting_time)
                    if bool_online:
                        self.pauseOption()
                        if not self.is_running:
                            # 后续测试全部取消
                            self.cancelAllTest()
                            # self.isTest = False
                            # self.CPU_isTest = [False for x in range(self.testNum)]
                            # self.testSign = False
                            break
                        self.result_signal.emit(f'总线初始化成功！' + self.HORIZONTAL_LINE)
                        break
                    else:
                        self.pauseOption()
                        if not self.is_running:
                            # 后续测试全部取消
                            self.cancelAllTest()
                            # self.isTest = False
                            # self.CPU_isTest = [False for x in range(self.testNum)]
                            # self.testSign = False
                            break
                        if eID != 0:
                            self.result_signal.emit(f'错误：未发现{module_array[eID-1]}' + self.HORIZONTAL_LINE)
                        self.result_signal.emit(f'错误：总线初始化失败！再次尝试初始化。' + self.HORIZONTAL_LINE)

                self.result_signal.emit('模块在线检测结束！' + self.HORIZONTAL_LINE)

            except:
                self.pauseOption()
                if not self.is_running:
                    # 后续测试全部取消
                    self.cancelAllTest()
                    # self.isTest = False
                    # self.CPU_isTest = [False for x in range(self.testNum)]
                    # self.testSign = False
                self.messageBox_signal.emit(['错误提示', '总线初始化异常，请检查设备！'])
                # 后续测试全部取消
                self.cancelAllTest()
                # self.isTest = False
                # self.CPU_isTest = [False for x in range(self.testNum)]
                # self.testSign = False
                self.result_signal.emit('模块在线检测结束！' + self.HORIZONTAL_LINE)
            ############################################################检测项目############################################################
            for i in range(len(self.CPU_isTest)):
                if self.isCancelAllTest:
                    break
                if self.CPU_isTest[i]:
                    if i == 0:#外观检测
                        self.CPU_appearanceTest()
                        if self.isCancelAllTest:
                            break
                    elif i == 1:#型号检查
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        try:
                            self.CPU_typeTest(itemRow=i)
                        except:
                            self.showErrorInf()
                            self.cancelAllTest()
                        finally:
                            if self.isPassTypeTest:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 2:#SRAM
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        try:
                            self.CPU_SRAMTest(itemRow=i)
                        except:
                            self.showErrorInf()
                            self.cancelAllTest()
                        finally:
                            if self.isPassSRAM:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 3:#FLASH
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        try:
                            self.CPU_FLASHTest()
                        except:
                            self.showErrorInf()
                            self.cancelAllTest()
                        finally:
                            if self.isPassFLASH:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 4:#FPGA
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        try:
                            self.messageBox_signal(["操作提示","请插入U盘，并观察U盘指示灯是否点亮？"])
                            reply = self.result_queue.get()
                            if reply == QMessageBox.Yes:
                                self.CPU_FPGATest()
                            elif reply == QMessageBox.No:
                                self.messageBox_signal(['测试警告', '无法识别U盘，FPGA无法测试，是否进行后续测试？'])
                                self.isPassFPGA = False
                                reply = self.result_queue.get()
                                if reply == QMessageBox.Yes:
                                    pass
                                elif reply == QMessageBox.No:
                                    self.cancelAllTest()
                        except:
                            self.showErrorInf()
                            self.cancelAllTest()
                        finally:
                            if self.isPassFPGA:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 5:#拨杆测试
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        try:
                            self.messageBox_signal(["操作提示","请将拨杆拨至STOP位置。"])
                            reply = self.result_queue.get()
                            if reply == QMessageBox.Yes or QMessageBox.No:
                                self.CPU_LeverTest(0)
                            if not self.isCancelAllTest:
                                self.messageBox_signal(["操作提示", "请将拨杆拨至RUN位置。"])
                                if reply == QMessageBox.Yes or QMessageBox.No:
                                    self.CPU_LeverTest(1)
                        except:
                            self.showErrorInf()
                            self.cancelAllTest()
                        finally:
                            if self.isPassLever:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 6:#MFK按钮
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        try:
                            self.CPU_MFKTest(0)
                            if not self.isCancelAllTest:
                                self.messageBox_signal(["操作提示", "请长按MKF按钮不要松开。"])
                                reply = self.result_queue.get()
                                if reply == QMessageBox.Yes or QMessageBox.No:
                                    self.CPU_MFKTest(1)
                                self.messageBox_signal(["操作提示", "请松开MKF按钮。"])
                                reply = self.result_queue.get()
                                if reply == QMessageBox.Yes or QMessageBox.No:
                                    pass
                        except:
                            self.showErrorInf()
                            self.cancelAllTest()
                        finally:
                            if self.isPassMFK:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 7:#掉电保存
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        try:
                            self.CPU_powerOffSaveTest()
                        except:
                            self.showErrorInf('掉电保存测试')
                            self.cancelAllTest()
                        finally:
                            if self.isPassPowerOffSave:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break

                    elif i == 8:#RTC
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        try:
                            testStartTime = time.time()
                            self.CPU_RTCTest('write')
                            if not self.isCancelAllTest:
                                self.CPU_RTCTest('read')
                        except:
                            self.showErrorInf('RTC测试')
                            self.cancelAllTest()
                        finally:
                            if self.isPassRTC:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 9:#各指示灯
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        try:
                            w_transmitData = [0xAC, hex(7), 0x00, 0x09, 0x10, 0x00, 0x01]
                            r_transmitData = [0xAC, hex(6), 0x00, 0x09, 0x0E, 0x00]
                            test_transmitData = [0xAC, hex(5), 0x00, 0x09, 0x53]
                            normal_transmitData = [0xAC, hex(5), 0x00, 0x09, 0x50]
                            name_LED = ['RUN灯','ERROR灯','BAT_LOW灯','PRG灯','232灯','485灯','HOST灯','所有灯']
                            #进入灯测试模式
                            self.CPU_LEDTest(itemRow=i, transmitData=test_transmitData,LED_name='')
                            if not self.isCancelAllTest:
                                for j in range(8):
                                    if j != 7:
                                        w_transmitData[5] = hex(j)
                                    else:
                                        w_transmitData[5] = 0xAA
                                    self.CPU_LEDTest(itemRow=i,transmitData=w_transmitData,LED_name=name_LED[j])
                                    if not self.isCancelAllTest:
                                        if j != 7:
                                            r_transmitData[5] = hex(j)
                                        else:
                                            r_transmitData[5] = 0xAA
                                        self.CPU_LEDTest(itemRow=i, transmitData=r_transmitData,LED_name=name_LED[j])
                                    else:
                                        break
                        except:
                            self.showErrorInf('LED测试')
                            self.cancelAllTest()
                        finally:
                            # 退出灯测试模式
                            self.CPU_LEDTest(itemRow=i, transmitData=test_transmitData, LED_name='')
                            if self.isPassLED:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 10:#本体IN
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        r_transmitData = [0xAC, hex(7),0x00, 0x0A, 0x0E, 0x00, 0xAA]
                        try:
                            #先控制两块QN0016模块向CPU输入端口发送高电平
                            self.m_transmitData = [0xff,0xff,0x00,0x00,0x00,0x00,0x00,0x00]
                            CAN_option.transmitCANAddr(0x200 + self.CANAddr1, self.m_transmitData)
                            CAN_option.transmitCANAddr(0x200 + self.CANAddr2, self.m_transmitData)
                            #1.读输入端口
                            self.CPU_InTest(transmitData=r_transmitData)
                        except:
                            self.showErrorInf('本体IN测试')
                            self.cancelAllTest()
                        finally:
                            if self.isPassIn:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 11:#本体OUT
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        w_transmitData = [0xAC, hex(10), 0x00, 0x0A, 0x10, 0x01, 0xAA, hex(2), 0xff, 0xff]
                        try:
                            self.CPU_OutTest(transmitData=w_transmitData)
                        except:
                            self.showErrorInf('本体OUT测试')
                            self.cancelAllTest()
                        finally:
                            if self.isPassOut:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 12:#以太网
                        self.item_signal.emit([i, 1, 0, ''])
                        testStartTime = time.time()
                        defaultIP_0 = 192
                        defaultIP_1 = 168
                        defaultIP_2 = 1
                        defaultIP_3 = 66
                        defaultUDP_LSB = 1213
                        defaultUDP_MSB = 50000

                        r_transmitData_IP = [0xAC, hex(9), 0x00, 0x0B, 0x0E, hex(defaultIP_0),
                                          hex(defaultIP_1), hex(defaultIP_2),
                                          hex(defaultIP_3)]
                        r_transmitData_UDP = [0xAC, hex(7), 0x00, 0x0B, 0x0E, hex(defaultUDP_LSB),
                                              hex(defaultUDP_MSB)]
                        try:
                            self.CPU_ETHTest(transmitData=r_transmitData_IP, isPassSign = self.isPassETH_IP)
                            if not self.isCancelAllTest:
                                self.CPU_ETHTest(transmitData=r_transmitData_UDP, isPassSign = self.isPassETH_UDP)
                        except:
                            self.showErrorInf('本体ETH测试')
                            self.cancelAllTest()
                        finally:
                            if self.isPassETH:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 13:#RS-232C
                        self.item_signal.emit([i, 1, 0, ''])
                        testStartTime = time.time()
                        # 115200 57600 38400 19200 9600  4800
                        byte0_array = [0x00, 0x00, 0x00, 0x00, 0x80, 0xC0]
                        byte1_array = [0xC2, 0xE1, 0x96, 0x4B, 0x25, 0x12]
                        byte2_array = [0x01, 0x00, 0x00, 0x00, 0x00, 0x00]
                        byte3_array = [0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
                        w_transmitData = [0xAC, hex(10), 0x00, 0x0C, 0x10, 0x00,
                                             0x00, 0x00, 0x00, 0x00]
                        r_transmitData = [0xAC, hex(6), 0x00, 0x0C, 0x0E, 0x00]
                        try:
                            for i in range(6):
                                w_transmitData[6] = byte0_array[i]
                                w_transmitData[7] = byte1_array[i]
                                w_transmitData[8] = byte2_array[i]
                                w_transmitData[9] = byte3_array[i]
                                self.CPU_232Test(transmitData=w_transmitData)
                                if not self.isCancelAllTest:
                                    self.CPU_232Test(transmitData=r_transmitData)
                                else:
                                    break
                        except:
                            self.showErrorInf('本体232测试')
                            self.cancelAllTest()
                        finally:
                            if self.isPass232:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 14:#RS-485
                        self.item_signal.emit([i, 1, 0, ''])
                        testStartTime = time.time()
                        # 115200 57600 38400 19200 9600  4800
                        byte0_array = [0x00, 0x00, 0x00, 0x00, 0x80, 0xC0]
                        byte1_array = [0xC2, 0xE1, 0x96, 0x4B, 0x25, 0x12]
                        byte2_array = [0x01, 0x00, 0x00, 0x00, 0x00, 0x00]
                        byte3_array = [0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
                        w_transmitData = [0xAC, hex(10), 0x00, 0x0C, 0x10, 0x00,
                                          0x00, 0x00, 0x00, 0x00]
                        r_transmitData = [0xAC, hex(6), 0x00, 0x0C, 0x0E, 0x00]
                        try:
                            for i in range(6):
                                w_transmitData[6] = byte0_array[i]
                                w_transmitData[7] = byte1_array[i]
                                w_transmitData[8] = byte2_array[i]
                                w_transmitData[9] = byte3_array[i]
                                self.CPU_485Test(transmitData=w_transmitData)
                                if not self.isCancelAllTest:
                                    self.CPU_485Test(transmitData=r_transmitData)
                                else:
                                    break
                        except:
                            self.showErrorInf('本体485测试')
                            self.cancelAllTest()
                        finally:
                            if self.isPass485:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 15:#右扩CAN
                        self.item_signal.emit([i, 1, 0, ''])
                        testStartTime = time.time()

                        transmitData_reset = [0xAC, hex(7), 0x00, 0x0E, 0x10, 0x00, 0x00]
                        transmitData_config = [0xAC, hex(7), 0x00, 0x0E, 0x10, 0x01, 0x01]
                        try:
                            self.CPU_rightCANTest(transmitData=transmitData_reset)
                            if not self.isCancelAllTest:
                                transmitData_reset[6] = 0x01
                                self.CPU_rightCANTest(transmitData=transmitData_reset)
                                if not self.isCancelAllTest:
                                    self.CPU_rightCANTest(transmitData=transmitData_config)
                                    if not self.isCancelAllTest:
                                        transmitData_config[6] = 0x00
                                        self.CPU_rightCANTest(transmitData=transmitData_config)
                                    else:
                                        break
                                else:
                                    break
                            else:
                                break
                        except:
                            self.showErrorInf('本体右扩CAN测试')
                            self.cancelAllTest()
                        finally:
                            if self.isPassRightCAN:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 16:#MAC/三码写入
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        #写MAC
                        w_MAC = [0xAC, hex(19), 0x00, 0x03, 0x10, 0x01,hex(12),
                                  ord(self.module_MAC[0]), ord(self.module_MAC[1]),
                                  ord(self.module_MAC[2]), ord(self.module_MAC[3]),
                                  ord(self.module_MAC[4]), ord(self.module_MAC[5]),
                                  ord(self.module_MAC[6]), ord(self.module_MAC[7]),
                                  ord(self.module_MAC[8]), ord(self.module_MAC[9]),
                                  ord(self.module_MAC[10]), ord(self.module_MAC[11])]
                        #读MAC
                        r_MAC = [0xAC, hex(6), 0x00, 0x03, 0x0E, 0x01]

                        # 写SN
                        w_SN = [0xAC, hex(19), 0x00, 0x03, 0x10, 0x02, hex(12),
                                ord(self.module_sn[0]), ord(self.module_sn[1]),
                                ord(self.module_sn[2]), ord(self.module_sn[3]),
                                ord(self.module_sn[4]), ord(self.module_sn[5]),
                                ord(self.module_sn[6]), ord(self.module_sn[7]),
                                ord(self.module_sn[8]), ord(self.module_sn[9]),
                                ord(self.module_sn[10]), ord(self.module_sn[11])]
                        # 读SN
                        r_SN = [0xAC, hex(6), 0x00, 0x03, 0x0E, 0x02]

                        # 写PN
                        w_PN = [0xAC, hex(21), 0x00, 0x03, 0x10, 0x03, hex(14),
                                 ord(self.module_pn[0]), ord(self.module_pn[1]),
                                 ord(self.module_pn[2]), ord(self.module_pn[3]),
                                 ord(self.module_pn[4]), ord(self.module_pn[5]),
                                 ord(self.module_pn[6]), ord(self.module_pn[7]),
                                 ord(self.module_pn[8]), ord(self.module_pn[9]),
                                 ord(self.module_pn[10]), ord(self.module_pn[11]),
                                 ord(self.module_pn[12]), ord(self.module_pn[13])]
                        # 读PN
                        r_PN = [0xAC, hex(6), 0x00, 0x03, 0x0E, 0x03]

                        try:
                            #写MAC
                            self.CPU_configTest(transmitData=w_MAC, type='MAC')
                            if not self.isCancelAllTest:
                                # 读MAC
                                self.CPU_configTest(transmitData=r_MAC, type='MAC')
                                if not self.isCancelAllTest:
                                    #写SN
                                    self.CPU_configTest(transmitData=w_SN, type='SN')
                                    if not self.isCancelAllTest:
                                        # 读SN
                                        self.CPU_configTest(transmitData=r_SN, type='SN')
                                        if not self.isCancelAllTest:
                                            # 写PN
                                            self.CPU_configTest(transmitData=w_PN, type='PN')
                                            if not self.isCancelAllTest:
                                                # 读PN
                                                self.CPU_configTest(transmitData=r_PN, type='PN')
                        except:
                            self.showErrorInf()
                            self.cancelAllTest()
                        finally:
                            if self.isPassConfig:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 17:#MA0202
                        pass

            if self.isCancelAllTest:
                self.result_signal("后续测试已全部取消，测试结束。")

        except:
            self.messageBox_signal.emit(['错误提示', f'测试出现问题，请检查测试程序和测试设备！\n'
                                                     f'"ErrorInf:\n{traceback.format_exc()}"'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes or QMessageBox.No:
                self.showErrorInf('测试')
                self.allFinished_signal.emit()
                self.pass_signal.emit(False)
    #CPU外观检查
    def CPU_appearanceTest(self):
        appearanceStart_time = time.time()
        self.item_signal.emit([0, 1, 0, ''])
        self.messageBox_signal.emit(['外观检测', '请检查：\n（1）外壳字体是否清晰?\n（2）型号是否正确？\n（3）外壳是否破损有污渍？'])
        reply = self.result_queue.get()
        if reply == QMessageBox.Yes:
            self.isPassAppearance = True
            appearanceEnd_time = time.time()
            appearanceTest_time = round(appearanceEnd_time - appearanceStart_time, 2)
            self.item_signal.emit([0, 2, 1, appearanceTest_time])
        elif reply == QMessageBox.No:
            self.isPassAppearance = False
            appearanceEnd_time = time.time()
            appearanceTest_time = round(appearanceEnd_time - appearanceStart_time, 2)
            self.item_signal.emit([0, 2, 2, appearanceTest_time])

        self.testNum = self.testNum - 1

    #CPU型号检查
    def CPU_typeTest(self,itemRow):
        serial_transmitData = [0xAC,hex(6),0x00,0x00,0x0E,0x00]
        #打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=1000000, timeout=1)
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime)*1000>self.waiting_time:
                self.isPassTypeTest = False
                break
            # 发送数据
            typeC_serial.write(bytes(serial_transmitData))
            # 等待0.5s后接收数据
            time.sleep(0.5)
            serial_receiveData = []
            # 接收数组数据
            serial_receiveData = typeC_serial.read(40)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 8:
                if trueData[6] == 0x00:#接收成功:
                    if trueData[5] == 0x00:#点数
                        if trueData[7] == 0x40:
                            self.isPassTypeTest = True
                            break
                        else:
                            self.isPassTypeTest = False
                            break
                    elif trueData[5] == 0x01:#类型
                        if trueData[7] == 0x00:#NPN
                            self.isPassTypeTest = True
                            break
                        else:
                            self.isPassTypeTest = False
                            break
                elif trueData[6] == 0x01:#接收失败
                    self.isPassTypeTest = False
                    break
            else:
                continue
        if not self.isPassTypeTest:
            self.messageBox_signal(['测试警告', '设备型号测试不通过，后续测试中止！'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes or reply == QMessageBox.No:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    #SRAM测试
    def CPU_SRAMTest(self):
        serial_transmitData = [0xAC, hex(5), 0x00, 0x01, 0x53]
        # 打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=1000000, timeout=1)
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassSRAM = False
                break
            # 发送数据
            typeC_serial.write(bytes(serial_transmitData))
            # 等待0.5s后接收数据
            time.sleep(0.5)
            serial_receiveData = []
            # 接收数组数据
            serial_receiveData = typeC_serial.read(40)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 6:
                if trueData[5] == 0x00:  # 接收成功
                    self.isPassSRAM = True
                    break
                elif trueData[5] == 0x01:  # 接收失败
                    self.isPassSRAM = False
                    break
            else:
                continue
        if not self.isPassSRAM:
            self.messageBox_signal(['测试警告', 'SRAM测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            elif reply == QMessageBox.No:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    #FLASH测试
    def CPU_FLASHTest(self):
        serial_transmitData = [0xAC, hex(5), 0x00, 0x02, 0x53]
        # 打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=1000000, timeout=1)
        loopStartTime = time.time()

        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassFLASH = False
                break
            # 发送数据
            typeC_serial.write(bytes(serial_transmitData))
            # 等待0.5s后接收数据
            time.sleep(0.5)
            serial_receiveData = []
            # 接收数组数据
            serial_receiveData = typeC_serial.read(40)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 6:
                if trueData[5] == 0x00:  # 接收成功
                    self.isPassFLASH = True
                    break
                elif trueData[5] == 0x01:  # 接收失败
                    self.isPassFLASH = False
                    break
            else:
                continue
        if not self.isPassFLASH:
            self.messageBox_signal(['测试警告', 'FLASH测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            elif reply == QMessageBox.No:
                self.cancelAllTest()

        # 关闭串口
        typeC_serial.close()

    # FPGA测试
    def CPU_FPGATest(self):
        serial_transmitData0 = [0xAC, hex(6), 0x00, 0x04, 0x53, 0x00]
        serial_transmitData1 = [0xAC, hex(6), 0x00, 0x04, 0x53, 0x01]
        serial_transmitData2 = [0xAC, hex(6), 0x00, 0x04, 0x53, 0x02]
        serial_transmitData_array = [serial_transmitData0,
                                     serial_transmitData1,
                                     serial_transmitData2]
        # 打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=1000000, timeout=1)
        loopStartTime = time.time()
        for serial_transmitData in serial_transmitData_array:
            while True:
                if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                    self.isPassFPGA = False
                    break
                # 发送数据
                typeC_serial.write(bytes(serial_transmitData))
                # 等待0.5s后接收数据
                time.sleep(0.5)
                serial_receiveData = []
                # 接收数组数据
                serial_receiveData = typeC_serial.read(40)
                trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
                if isSendAgain:
                    continue
                if dataLen == 3:  # 指令出错
                    self.orderError(trueData[2])
                    break
                elif dataLen == 7:
                    if trueData[6] == 0x00:  # 成功
                        self.isPassFPGA = True
                        break
                    elif trueData[6] == 0x01:  # 失败
                        self.isPassFPGA = False
                        break
                else:
                    continue
            if not self.isPassFPGA:
                break
        if not self.isPassFPGA:
            self.messageBox_signal(['测试警告', 'FPGA测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            elif reply == QMessageBox.No:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    # 拨杆测试
    def CPU_LeverTest(self,state:int):
        serial_transmitData = [0xAC, hex(6), 0x00, 0x05, 0x0E, 0x00]
        # 打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=1000000, timeout=1)
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassLever = False
                break
            # 发送数据
            typeC_serial.write(bytes(serial_transmitData))
            # 等待0.5s后接收数据
            time.sleep(0.5)
            serial_receiveData = []
            # 接收数组数据
            serial_receiveData = typeC_serial.read(40)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 8: #成功
                if trueData[7] == hex(state):  # 测试通过
                    self.isPassLever = True
                    break
                else:  # 测试不通过
                    self.isPassLever = False
                    break
            elif dataLen == 7:  # 失败
                self.isPassLever = False
                break
            else:
                continue
        if not self.isPassLever:
            self.messageBox_signal(['测试警告', '拨杆测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            elif reply == QMessageBox.No:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    # MFK测试
    def CPU_MFKTest(self,state:int):
        serial_transmitData = [0xAC, hex(6), 0x00, 0x06, 0x0E, 0x00]
        # 打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=1000000, timeout=1)
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassLever = False
                break
            # 发送数据
            typeC_serial.write(bytes(serial_transmitData))
            # 等待0.5s后接收数据
            time.sleep(0.5)
            serial_receiveData = []
            # 接收数组数据
            serial_receiveData = typeC_serial.read(40)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 8:#成功
                if trueData[7] == hex(state):#测试通过
                    self.isPassLever = True
                    break
                else:#测试不通过
                    self.isPassLever = False
                    break
            elif dataLen == 7:  # 失败
                self.isPassMFK = False
                break
            else:
                continue
        if not self.isPassMFK:
            self.messageBox_signal(['测试警告', 'MFK测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            elif reply == QMessageBox.No:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    # RTC测试
    def CPU_RTCTest(self,wr:str):
        serial_transmitData=[]
        if wr == 'write':
            now = datetime.datetime.now()
            serial_transmitData = [0xAC, hex(13), 0x00, 0x07, 0x10, 0x00,
                                        hex(now.year-2000), hex(now.month), hex(now.day),
                                        hex(now.hour), hex(now.minute), hex(now.second),
                                        hex(now.weekday())]
        elif wr == 'read':
            serial_transmitData = [0xAC, hex(6), 0x00, 0x07, 0x0E, 0x00]

        # 打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=1000000, timeout=1)
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassRTC = False
                break
            # 发送数据
            typeC_serial.write(bytes(serial_transmitData))
            # 等待0.5s后接收数据
            time.sleep(0.5)
            serial_receiveData = []
            # 接收数组数据
            serial_receiveData = typeC_serial.read(40)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 7: #写
                if trueData[6] == 0x00:  # 写成功
                    self.isPassRTC = True
                    break
                elif trueData[6] == 0x01:  # 写失败
                    self.isPassRTC = True
                    break
            elif dataLen == 14:#读
                if trueData[6] == 0x00:  # 读成功
                    now = datetime.datetime.now()
                    if (now.year-2000) == trueData[7] and now.month == trueData[8] and\
                        now.day == trueData[9] and now.hour == trueData[10] and\
                        now.weekday() == trueData[13]:
                        if (now.hour-trueData[10]) *3600+(now.minute-trueData[11])*60\
                                +(now.second-trueData[12])<30:
                            self.isPassRTC = True
                            break
                        else:
                            self.isPassRTC = False
                            break
                    else:
                        self.isPassRTC = False
                        break
                elif trueData[6] == 0x01:  # 读失败
                    break
            else:
                continue
        if not self.isPassRTC:
            self.messageBox_signal(['测试警告', 'RTC测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            elif reply == QMessageBox.No:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    # 掉电保存测试
    def CPU_powerOffSaveTest(self):
        serial_transmitData0 = [0xAC, hex(6), 0x00, 0x08, 0x53, 0x00]
        serial_transmitData1 = [0xAC, hex(6), 0x00, 0x08, 0x53, 0x01]
        serial_transmitData2 = [0xAC, hex(6), 0x00, 0x08, 0x53, 0xFF]
        serial_transmitData_array = [serial_transmitData0,
                                     serial_transmitData1,serial_transmitData2]
        # 打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=1000000, timeout=1)
        loopStartTime = time.time()
        for serial_transmitData in serial_transmitData_array:
            while True:
                if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                    self.isPassPowerOffSave = False
                    break
                # 发送数据
                typeC_serial.write(bytes(serial_transmitData))
                # 等待0.5s后接收数据
                time.sleep(0.5)
                serial_receiveData = []
                # 接收数组数据
                serial_receiveData = typeC_serial.read(40)
                trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
                if isSendAgain:
                    continue
                if dataLen == 3:  # 指令出错
                    self.orderError(trueData[2])
                    break
                elif dataLen == 7:
                    if trueData[6] == 0x00:  # 成功
                        self.isPassPowerOffSave = True
                        break
                    elif trueData[6] == 0x01:  # 失败
                        self.isPassPowerOffSave = False
                        break
                else:
                    continue
            if not self.isPassPowerOffSave:
                break
        if not self.isPassPowerOffSave:
            self.messageBox_signal(['测试警告', '掉电保存测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            elif reply == QMessageBox.No:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    # 各指示灯测试
    def CPU_LEDTest(self, itemRow:int, transmitData:list,LED_name:str):
        # 打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=1000000, timeout=1)
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassLED = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            # 等待0.5s后接收数据
            time.sleep(0.5)
            serial_receiveData = []
            # 接收数组数据
            serial_receiveData = typeC_serial.read(40)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                break
            if dataLen == 6:
                if trueData[5] == 0x00:  # 进入（退出）测试模式成功
                    break
                elif trueData[5] == 0x01:  # 进入（退出）测试模式失败
                    self.isPassLED = False
                    break
            elif dataLen == 7: #写
                if trueData[6] == 0x00:  # 写入成功
                    self.isPassLED = True
                elif trueData[6] == 0x01:  # 写入失败
                    self.isPassLED = False
                break
            elif dataLen == 8: #读
                if trueData[7] == 0x00:#灯熄灭
                    self.messageBox_signal(['操作提示',f'{LED_name}是否熄灭？'])
                    reply = self.result_queue.get()
                    if reply == QMessageBox.Yes:
                        self.isPassLED = True
                    elif reply == QMessageBox.No:
                        self.isPassLED = False
                    break
                elif trueData[7] == 0x01:#灯点亮
                    self.messageBox_signal(['操作提示',f'{LED_name}是否点亮？'])
                    reply = self.result_queue.get()
                    if reply == QMessageBox.Yes:
                        self.isPassLED = True
                    elif reply == QMessageBox.No:
                        self.isPassLED = False
                    break
            elif dataLen == 9:#控制所有灯，暂不用这个模式
                pass
            else:
                continue
        if not self.isPassLED:
            self.messageBox_signal(['测试警告', 'LED测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            elif reply == QMessageBox.No:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    def CPU_InTest(self, transmitData:list):
        # 打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=1000000, timeout=1)
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassIn = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            # 等待0.5s后接收数据
            time.sleep(0.5)
            serial_receiveData = []
            # 接收数组数据
            serial_receiveData = typeC_serial.read(40)
            trueData,dataLen,isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                continue
            if dataLen == 3:#指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 12:#读所有输入通道
                if trueData[7] == 0x00:  # 读所有输入通道成功
                    if trueData[9] == 0xff and trueData[10] == 0xff and trueData[11] == 0xff:
                        self.messageBox_signal(['操作提示','输入通道指示灯是否全部点亮？'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.Yes:
                            self.isPassIn = True
                        elif reply == QMessageBox.No:
                            self.isPassIn = False
                        break
                    else:
                        self.isPassIn = False
                        break
                elif trueData[7] == 0x01:  # 读所有输入通道失败
                    self.isPassIn = False
                    break

        if not self.isPassIn:
            self.messageBox_signal(['测试警告', 'LED测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            elif reply == QMessageBox.No:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    #本体OUT测试
    def CPU_OutTest(self, transmitData:list):
        # 打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=1000000, timeout=1)
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassOut = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            # 等待0.5s后接收数据
            time.sleep(0.5)
            serial_receiveData = []
            # 接收数组数据
            serial_receiveData = typeC_serial.read(40)
            trueData,dataLen,isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                continue
            if dataLen == 3:#指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 8:#读所有输出通道
                if trueData[7] == 0x00:  # 读所有输出通道成功
                    bool_receive, self.m_can_obj = CAN_option.receiveCANbyID((0x180 + self.CANAddr_DI), 2000)
                    self.m_receiveData = self.m_can_obj.Data
                    if bool_receive == False:
                        self.messageBox_signal(['警告','ET1600接收数据超时，请检查ET1600接线！'])
                        self.isPassOut = False
                        break
                    elif self.m_receiveData[0] == 0xff and self.m_receiveData[1] == 0xff :
                        self.messageBox_signal(['操作提示','ET1600的通道指示灯是否全部点亮？'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.Yes:
                            self.isPassOut = True
                        elif reply == QMessageBox.No:
                            self.isPassOut = False
                        break
                elif trueData[7] == 0x01:  # 读所有输出通道失败
                    self.isPassOut = False
                    break
            else:
                continue

        if not self.isPassOut:
            self.messageBox_signal(['测试警告', '本体OUT测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            elif reply == QMessageBox.No:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    # 本体ETH测试
    def CPU_ETHTest(self, transmitData: list,isPassSign:bool):
        # 打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=1000000, timeout=1)
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassETH = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            # 等待0.5s后接收数据
            time.sleep(0.5)
            serial_receiveData = []
            # 接收数组数据
            serial_receiveData = typeC_serial.read(40)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 11:  # 读IP
                if trueData[6] == 0x00:  # 读IP成功
                    if trueData[7] == transmitData[5] and trueData[8] == transmitData[6] \
                            and trueData[9] == transmitData[7] and trueData[10] == transmitData[8]:
                        isPassSign = True
                        break
                    else:
                        isPassSign = False
                        break
                elif trueData[6] == 0x01:  # 读IP失败
                    isPassSign = False
                    break
            elif dataLen == 9:  # 读UDP
                if trueData[6] == 0x00:  # 读UDP成功
                    if trueData[7] == transmitData[5] and trueData[8] == transmitData[6]:
                        isPassSign = True
                        break
                    else:
                        isPassSign = False
                        break
                elif trueData[6] == 0x01:  # 读UDP失败
                    isPassSign = False
                    break
            else:
                continue

        if not self.isPassETH_IP and self.isPassETH_UDP:
            self.messageBox_signal(['测试警告', '本体ETH IP测试不通过，是否进行后续测试？'])
            self.isPassETH = False
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            elif reply == QMessageBox.No:
                self.cancelAllTest()
        elif (not self.isPassETH_UDP and self.isPassETH_IP) or (not self.isPassETH_UDP and not self.isPassETH_IP):
            self.isPassETH = False
            self.messageBox_signal(['测试警告', '本体ETH UDP测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            elif reply == QMessageBox.No:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    # 本体232测试
    def CPU_232Test(self, transmitData: list):
        # 打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=1000000, timeout=1)
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPass232 = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            # 等待0.5s后接收数据
            time.sleep(0.5)
            serial_receiveData = []
            # 接收数组数据
            serial_receiveData = typeC_serial.read(40)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 7:  # 写
                if trueData[6] == 0x00:  # 写成功
                    self.isPass232 = True
                    break
                elif trueData[6] == 0x01:  # 写/读失败
                    self.isPass232 = False
                    break
            elif dataLen == 11:  # 读
                if trueData[6] == 0x00:  # 读成功
                    if trueData[7] == transmitData[6] and trueData[8] == transmitData[7]\
                        and trueData[9] == transmitData[8] and trueData[10] == transmitData[9]:
                        self.isPass232 = True
                    else:
                        self.isPass232 = False
                    break
                elif trueData[6] == 0x01:  # 读失败
                    self.isPass232 = False
                    break
            else:
                continue
        if not self.isPass232:
            self.messageBox_signal(['测试警告', '本体232测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            elif reply == QMessageBox.No:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    # 本体485测试
    def CPU_485Test(self, transmitData: list):
        # 打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=1000000, timeout=1)
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPass485 = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            # 等待0.5s后接收数据
            time.sleep(0.5)
            serial_receiveData = []
            # 接收数组数据
            serial_receiveData = typeC_serial.read(40)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 7:  # 写
                if trueData[6] == 0x00:  # 写成功
                    self.isPass485 = True
                    break
                elif trueData[6] == 0x01:  # 写/读失败
                    self.isPass485 = False
                    break
            elif dataLen == 11:  # 读
                if trueData[6] == 0x00:  # 读成功
                    if trueData[7] == transmitData[6] and trueData[8] == transmitData[7] \
                            and trueData[9] == transmitData[8] and trueData[10] == transmitData[9]:
                        self.isPass485 = True
                    else:
                        self.isPass485 = False
                    break
                elif trueData[6] == 0x01:  # 读失败
                    self.isPass485 = False
                    break
            else:
                continue
        if not self.isPass485:
            self.messageBox_signal(['测试警告', '本体485测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            elif reply == QMessageBox.No:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    # 本体右扩CAN测试
    def CPU_rightCANTest(self, transmitData: list):
        # 打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=1000000, timeout=1)
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassRightCAN = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            # 等待0.5s后接收数据
            time.sleep(0.5)
            serial_receiveData = []
            # 接收数组数据
            serial_receiveData = typeC_serial.read(40)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 7:  # 写
                if trueData[6] == 0x00:  # 写成功
                    if transmitData[5] == 0x00 and transmitData[6] == 0x00:#复位灯亮
                        self.messageBox_signal(['操作提示','右扩CAN的RESET灯是否点亮?'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.Yes:
                            self.isPassRightCAN = True
                        elif reply == QMessageBox.No:
                            self.isPassRightCAN = False
                    elif transmitData[5] == 0x00 and transmitData[6] == 0x01:#复位灯灭
                        self.messageBox_signal(['操作提示','右扩CAN的RESET灯是否熄灭?'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.Yes:
                            self.isPassRightCAN = True
                        elif reply == QMessageBox.No:
                            self.isPassRightCAN =False
                    elif transmitData[5] == 0x01 and transmitData[6] == 0x01:#使能灯亮
                        self.messageBox_signal(['操作提示','右扩CAN的使能灯是否点亮?'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.Yes:
                            self.isPassRightCAN = True
                        elif reply == QMessageBox.No:
                            self.isPassRightCAN =False
                    elif transmitData[5] == 0x01 and transmitData[6] == 0x00:#使能灯灭
                        self.messageBox_signal(['操作提示','右扩CAN的使能灯是否熄灭?'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.Yes:
                            self.isPassRightCAN = True
                        elif reply == QMessageBox.No:
                            self.isPassRightCAN =False
                    break
                elif trueData[6] == 0x01:  # 写失败
                    self.isPassRightCAN = False
                    break
            else:
                continue
        if not self.isPassRightCAN:
            self.messageBox_signal(['测试警告', '本体右扩CAN测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            elif reply == QMessageBox.No:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    #MAC地址和3码
    def CPU_configTest(self, transmitData: list, type:str):
        # 打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=1000000, timeout=1)
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassConfig = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            # 等待0.5s后接收数据
            time.sleep(0.5)
            serial_receiveData = []
            # 接收数组数据
            serial_receiveData = typeC_serial.read(40)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 7:  # 写
                if trueData[6] == 0x00:  # 写成功
                    self.isPassConfig = True
                    break
                elif trueData[6] == 0x01:  # 写/读失败
                    self.isPassConfig = False
                    break
            elif dataLen == 20 and type == 'MAC':  # 读MAC
                if trueData[6] == 0x00:  # 读成功
                    if trueData[7] == 12:
                        for MACNum in range(12):
                            if not (trueData[MACNum+8] == ord(self.module_MAC[MACNum])):
                                self.isPassConfig = False
                                break
                            else:
                                self.isPassConfig = True
                    else:
                        self.isPassConfig = False
                    break
                elif trueData[6] == 0x01:  # 读失败
                    self.isPassConfig = False
                    break
            elif dataLen == 20 and type == 'SN':  # 读SN
                if trueData[6] == 0x00:  # 读成功
                    if trueData[7] == 12:
                        for SNNum in range(12):
                            if not (trueData[SNNum+8] == ord(self.module_pn[SNNum])):
                                self.isPassConfig = False
                                break
                            else:
                                self.isPassConfig = True
                    else:
                        self.isPassConfig = False
                    break
                elif trueData[6] == 0x01:  # 读失败
                    self.isPassConfig = False
                    break
            elif dataLen == 22 and type == 'PN':  # 读PN
                if trueData[6] == 0x00:  # 读成功
                    if trueData[7] == 12:
                        for PNNum in range(14):
                            if not (trueData[PNNum+8] == ord(self.module_pn[PNNum])):
                                self.isPassConfig = False
                                break
                            else:
                                self.isPassConfig = True
                    else:
                        self.isPassConfig = False
                    break
                elif trueData[6] == 0x01:  # 读失败
                    self.isPassConfig = False
                    break
            else:
                continue
        if not self.isPassConfig:
            self.messageBox_signal(['测试警告', '本体Config测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            elif reply == QMessageBox.No:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    #判断接收的数据是否正确
    def dataJudgement(self,serial_receiveData:list):
        isSendAgain = False
        trueData = [ ]
        dataLen = 0
        # 数据头位置
        startNum = 0
        # 确定数据头
        for i in range(len(serial_receiveData)):
            if serial_receiveData[i] == 0xCA:
                startNum = i
                isSendAgain = False
                break
            if i == len(serial_receiveData) - 1 and serial_receiveData[i] != 0xCA:
                isSendAgain = True
                return trueData,dataLen,isSendAgain
        dataLen = int(serial_receiveData[startNum + 1])

        # 验证数据接收是否正确
        dataSum = 0
        for i in range(dataLen):
            dataSum += serial_receiveData[startNum + i]
        ckeckSum = -int(dataSum)
        if ckeckSum == int(serial_receiveData[startNum + dataLen]):
            trueData = serial_receiveData[startNum:startNum + dataLen]
        return trueData,dataLen,isSendAgain

    def orderError(self,errCode):
        errorCode = {0x80: '消息长度错误', 0x81: '设备错误', 0x82: '测试类型错误',
                     0x83: '操作不支持', 0x84: '参数错误'}
        self.showErrorInf(f'指令发送错误：{errCode}，{errorCode[errCode]}\n'
                          f'取消剩余测试。\n')
        self.cancelAllTest()

    def changeTabItem(self,testStartTime,row,state,result):
        """
        :param testStartTime:
        :param row: 进行操作的单元行
        :param state: 测试状态 ['无需检测','正在检测','检测完成','等待检测']
        :param result: 测试结果 col2 = ['','通过','未通过']
        :return:
        """
        testEndTime = time.time()
        testAllTime = round(testEndTime - testStartTime, 2)
        self.item_signal.emit([row, state, result, testAllTime])

    #取消后续所有测试
    def cancelAllTest(self):
        self.isCancelAllTest = True

    #数组元素归零
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
                    break

    def stop_work(self):
        self.resume_work()
        self.is_running = False
        self.setAOChOutCalibrate()
        self.label_signal.emit(['fail', '测试停止'])

    def showErrorInf(self,Inf):
        self.showInf(f"{Inf}程序出现问题")
        # 捕获异常并输出详细的错误信息
        self.showInf(f"ErrorInf:\n{traceback.format_exc()}")

