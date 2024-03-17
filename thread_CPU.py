# -*- coding: utf-8 -*-

import time

from PyQt5.QtCore import QThread, pyqtSignal

import CAN_option
from main_logic import *
from CAN_option import *
import otherOption
import logging
import struct
import threading

class CPUThread(QObject):
    result_signal = pyqtSignal(str)
    item_signal = pyqtSignal(list)
    pass_signal = pyqtSignal(bool)

    messageBox_signal = pyqtSignal(list)
    pic_messageBox_signal = pyqtSignal(list)
    messageBox_oneButton_signal = pyqtSignal(list)
    # excel_signal = pyqtSignal(list)
    allFinished_signal = pyqtSignal()
    label_signal = pyqtSignal(list)
    saveExcel_signal = pyqtSignal(list)
    print_signal = pyqtSignal(list)

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
    isPassAppearance = True
    isPassTypeTest = True
    isPassSRAM = True
    isPassFLASH = True
    isPassConfig = True
    isPassFPGA = True
    isPassLever = True
    isPassMFK = True
    isPassRTC = True
    isPassPowerOffSave = True
    isPassLED = True
    isPassIn = True
    isPassOut =True
    isPassETH  = True
    isPass232 = True
    isPass485 = True
    isPassRightCAN = True
    isPassOp = True

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
        # self.testNum = 17
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
        self.serialPort_power = inf_CPUlist[3][4]
        self.serialPort_Op232 = inf_CPUlist[3][5]
        self.serialPort_Op485 = inf_CPUlist[3][6]

        # 获取CPU检测信息
        self.CPU_isTest = inf_CPUlist[4][0]
        self.CPU_testItem = [False for x in range(len(self.CPU_isTest))]
        self.CPU_passItem = [False for x in range(len(self.CPU_isTest))]
        self.isExcel = inf_CPUlist[4][1]

        #信号等待时间(ms)
        self.waiting_time = 2000

        self.isCancelAllTest = False


        self.errorNum = 0
        self.errorInf = ''
        self.isPassAll = True
        self.pause_num = 1
        #CPU输入通道错误信息
        self.inErrorInf = '输入通道'
        # CPU输出通道错误信息
        self.outErrorInf = '输出通道'

        self.current_dir = inf_CPUlist[5]
        self.testType = inf_CPUlist[6]

        self.pic_dir = current_dir + '/pic'
        self.CPU_pic_dir = self.pic_dir + '/CPU'
        # #初始化表格状态
        # for i in range(len(self.CPU_isTest)):
        #     self.result_signal.emit(f"self.CPU_isTest[{i}]:{self.CPU_isTest[i]}\n")
        #     print(f"self.CPU_isTest[{i}]:{self.CPU_isTest[i]}\n")
        #     if self.CPU_isTest[i]:
        #         self.item_signal.emit([i, 3, 0, ''])
        #     else:
        #         self.item_signal.emit([i, 0, 0, ''])

    def CPU_start(self):
        thread_signal = self.CPUOption()
        # if not thread_signal:
        #     self.label_signal.emit(['fail', '未通过'])

    def CPUOption(self):
        # self.isExcel = True
        #测试是否成功标志
        # self.testSign = True
        try:
            ############################################################总线初始化############################################################
            # try:
            #     time_online = time.time()
            #     while True:
            #         QApplication.processEvents()
            #         if (time.time() - time_online) * 1000 > self.waiting_time:
            #             self.pauseOption()
            #             if not self.is_running:
            #                 # 后续测试全部取消
            #                 self.cancelAllTest()
            #                 # self.isTest = False
            #                 # self.CPU_isTest = [False for x in range(self.testNum)]
            #                 # self.testSign = False
            #                 break
            #             self.result_signal.emit(f'错误：总线初始化超时！' + self.HORIZONTAL_LINE)
            #             self.messageBox_signal.emit(['错误提示', '总线初始化超时！请检查CAN分析仪或各设备是否正确连接！'])
            #             # 后续测试全部取消
            #             self.cancelAllTest()
            #             # self.isTest = False
            #             # self.CPU_isTest = [False for x in range(self.testNum)]
            #             # self.testSign = False
            #             break
            #         CANAddr_array = [self.CANAddr1,self.CANAddr2,self.CANAddr3,self.CANAddr4,self.CANAddr5]
            #         module_array = [self.module_1, self.module_2, self.module_3, self.module_4, self.module_5]
            #         bool_online, eID = otherOption.isModulesOnline(CANAddr_array,module_array, self.waiting_time)
            #         if bool_online:
            #             self.pauseOption()
            #             if not self.is_running:
            #                 # 后续测试全部取消
            #                 self.cancelAllTest()
            #                 # self.isTest = False
            #                 # self.CPU_isTest = [False for x in range(self.testNum)]
            #                 # self.testSign = False
            #                 break
            #             self.result_signal.emit(f'总线初始化成功！' + self.HORIZONTAL_LINE)
            #             break
            #         else:
            #             self.pauseOption()
            #             if not self.is_running:
            #                 # 后续测试全部取消
            #                 self.cancelAllTest()
            #                 # self.isTest = False
            #                 # self.CPU_isTest = [False for x in range(self.testNum)]
            #                 # self.testSign = False
            #                 break
            #             if eID != 0:
            #                 self.result_signal.emit(f'错误：未发现{module_array[eID-1]}' + self.HORIZONTAL_LINE)
            #             self.result_signal.emit(f'错误：总线初始化失败！再次尝试初始化。' + self.HORIZONTAL_LINE)
            #
            #     self.result_signal.emit('模块在线检测结束！' + self.HORIZONTAL_LINE)
            #
            # except:
            #     self.pauseOption()
            #     if not self.is_running:
            #         # 后续测试全部取消
            #         self.cancelAllTest()
            #         # self.isTest = False
            #         # self.CPU_isTest = [False for x in range(self.testNum)]
            #         # self.testSign = False
            #     self.messageBox_signal.emit(['错误提示', '总线初始化异常，请检查设备！'])
            #     # 后续测试全部取消
            #     self.cancelAllTest()
            #     # self.isTest = False
            #     # self.CPU_isTest = [False for x in range(self.testNum)]
            #     # self.testSign = False
            #     self.result_signal.emit('模块在线检测结束！' + self.HORIZONTAL_LINE)
            ############################################################检测项目############################################################
            #先测试一下远程CAN和DIP_SW
            # self.l_distanceCAN()
            # self.DIP_SW()
            for i in range(len(self.CPU_isTest)):
                if self.isCancelAllTest:
                    break
                if self.CPU_isTest[i]:
                    if i == 0:#外观检测
                        self.moveToRow_signal.emit([i, 0])
                        self.pauseOption()
                        if not self.is_running:
                            self.cancelAllTest()
                            break
                        self.result_signal.emit(f'----------------外观检测----------------')
                        self.CPU_appearanceTest()
                        if self.isCancelAllTest:
                            break
                        self.CPU_testItem[i] = True
                        self.isPassAll &= self.isPassAppearance
                        self.CPU_passItem[i] = self.isPassAppearance
                    elif i == 1:#型号检查
                        self.moveToRow_signal.emit([i, 0])
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.pauseOption()
                        if not self.is_running:
                            self.cancelAllTest()
                            break
                        self.result_signal.emit(f'----------------型号测试----------------')
                        try:
                            #读点数:
                            serial_transmitData = [0xAC, 6, 0x00, 0x00, 0x0E, 0x00]
                            self.CPU_typeTest(serial_transmitData=[0xAC,6,0x00,0x00,0x0E,0x00,
                                                                   self.getCheckNum(serial_transmitData)])
                            if not self.isCancelAllTest:
                                # 读类型:
                                serial_transmitData = [0xAC, 6, 0x00, 0x00, 0x0E, 0x01]
                                self.CPU_typeTest(serial_transmitData=[0xAC, 6, 0x00, 0x00, 0x0E, 0x01,
                                                                       self.getCheckNum(serial_transmitData)])
                        except:
                            self.isPassTypeTest &= False
                            self.showErrorInf('型号检查')
                            self.cancelAllTest()
                        finally:
                            self.CPU_passItem[i] = self.isPassTypeTest
                            self.CPU_testItem[i] = True
                            self.isPassAll &= self.isPassTypeTest
                            self.testNum = self.testNum - 1
                            if self.isPassTypeTest:
                                self.result_signal.emit("型号检查通过"+self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("型号检查未通过"+self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 2:#SRAM
                        self.moveToRow_signal.emit([i, 0])
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------SRAM测试----------------')
                        try:
                            serial_transmitData = [0xAC, 5, 0x00, 0x01, 0x53]
                            self.CPU_SRAMTest(serial_transmitData=[0xAC, 5, 0x00, 0x01, 0x53,
                                                                 self.getCheckNum(serial_transmitData)])
                        except:
                            self.isPassSRAM &= False
                            self.showErrorInf('SRAM')
                            self.cancelAllTest()
                        finally:
                            self.CPU_passItem[i] = self.isPassSRAM
                            self.CPU_testItem[i] = True
                            self.isPassAll &= self.isPassSRAM
                            self.testNum = self.testNum - 1
                            if self.isPassSRAM:
                                self.result_signal.emit("SRAM测试通过"+self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("SRAM测试未通过"+self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 3:#FLASH
                        self.moveToRow_signal.emit([i, 0])
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------FLASH测试----------------')
                        try:
                            self.CPU_FLASHTest(serial_transmitData = [0xAC, 5, 0x00, 0x02, 0x53,
                                                                      self.getCheckNum([0xAC, 5, 0x00, 0x02, 0x53])])
                        except:
                            self.showErrorInf('FLASH')
                            self.cancelAllTest()
                        finally:
                            self.CPU_testItem[i] = True
                            self.CPU_passItem[i] = self.isPassFLASH
                            self.isPassAll &= self.isPassFLASH
                            self.testNum = self.testNum - 1
                            if self.isPassFLASH:
                                self.result_signal.emit("FLASH测试通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("FLASH测试未通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 4:#拨杆测试
                        self.moveToRow_signal.emit([i, 0])
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------拨杆测试----------------')
                        try:
                            self.messageBox_signal.emit(["操作提示","请将拨杆拨至STOP位置（下）。"])
                            reply = self.result_queue.get()
                            if reply == QMessageBox.AcceptRole or QMessageBox.RejectRole:
                                self.CPU_LeverTest(0,serial_transmitData = [0xAC, 6, 0x00, 0x05, 0x0E, 0x00,
                                                                self.getCheckNum([0xAC, 6, 0x00, 0x05, 0x0E, 0x00])])
                            if not self.isCancelAllTest:
                                self.messageBox_signal.emit(["操作提示", "请将拨杆拨至RUN位置（上）。"])
                                reply = self.result_queue.get()
                                if reply == QMessageBox.AcceptRole or QMessageBox.RejectRole:
                                    self.CPU_LeverTest(1,serial_transmitData = [0xAC, 6, 0x00, 0x05, 0x0E, 0x00,
                                                                self.getCheckNum([0xAC, 6, 0x00, 0x05, 0x0E, 0x00])])
                        except:
                            self.isPassLever = False
                            self.showErrorInf('拨杆测试')
                            self.cancelAllTest()
                        finally:
                            self.CPU_passItem[i] = self.isPassLever
                            self.CPU_testItem[i] = True
                            self.isPassAll &= self.isPassLever
                            self.testNum = self.testNum - 1
                            if self.isPassLever:
                                self.result_signal.emit("拨杆测试通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("拨杆测试未通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 5:#MFK按钮
                        self.moveToRow_signal.emit([i, 0])
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------MFK测试----------------')
                        try:
                            self.CPU_MFKTest(0,serial_transmitData = [0xAC, 6, 0x00, 0x06, 0x0E, 0x00,
                                                                  self.getCheckNum([0xAC, 6, 0x00, 0x06, 0x0E, 0x00])])
                            if not self.isCancelAllTest:
                                self.messageBox_signal.emit(["操作提示", "请长按MFK按钮不要松开。\n（按住同时点确定）"])
                                reply = self.result_queue.get()
                                if reply == QMessageBox.AcceptRole or QMessageBox.RejectRole:
                                    self.CPU_MFKTest(1,serial_transmitData = [0xAC, 6, 0x00, 0x06, 0x0E, 0x00,
                                                                  self.getCheckNum([0xAC, 6, 0x00, 0x06, 0x0E, 0x00])])
                                self.messageBox_signal.emit(["操作提示", "请松开MFK按钮。\n（请操作后点确定）"])
                                reply = self.result_queue.get()
                                if reply == QMessageBox.AcceptRole or QMessageBox.RejectRole:
                                    pass
                        except:
                            self.isPassMFK = False
                            self.showErrorInf('MFK按钮')
                            self.cancelAllTest()
                        finally:
                            self.CPU_passItem[i] = self.isPassMFK
                            self.CPU_testItem[i] = True
                            self.isPassAll &= self.isPassMFK
                            self.testNum = self.testNum - 1
                            if self.isPassMFK:
                                self.result_signal.emit("MFK测试通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("MFK测试未通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            image_485boma = self.CPU_pic_dir+'/485拨码.png'
                            self.pic_messageBox_signal.emit(['操作提示','请将如图所示的485拨码拨至左侧。',image_485boma])
                            reply = self.result_queue.get()
                            if reply == QMessageBox.AcceptRole or reply == QMessageBox.RejectRole:
                                if self.isCancelAllTest:
                                    break
                    elif i == 6:#掉电保存
                        self.moveToRow_signal.emit([i, 0])
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------掉电保存测试----------------')
                        try:
                            serial_transmitData0 = [0xAC, 6, 0x00, 0x08, 0x53, 0x00,
                                                    self.getCheckNum([0xAC, 6, 0x00, 0x08, 0x53, 0x00])]
                            serial_transmitData1 = [0xAC, 6, 0x00, 0x08, 0x53, 0x01,
                                                    self.getCheckNum([0xAC, 6, 0x00, 0x08, 0x53, 0x01])]
                            serial_transmitData2 = [0xAC, 6, 0x00, 0x08, 0x53, 0xFF,
                                                    self.getCheckNum([0xAC, 6, 0x00, 0x08, 0x53, 0xFF])]
                            serial_transmitData_array = [serial_transmitData0,
                                                         serial_transmitData1, serial_transmitData2]
                            #数据写入
                            self.CPU_powerOffSaveTest(serial_transmitData=serial_transmitData0)
                            if not self.isCancelAllTest:
                                if self.CPU_isTest[7]:
                                    self.CPU_RTCTest('write')
                                    if not self.isPassRTC:
                                        self.CPU_isTest[7] = False
                                        self.result_signal.emit('RTC写入失败，取消后续RTC测试！\n\n')
                                # 可编程电源开断电
                                power_off = [0x01, 0x06, 0x00, 0x01, 0x00, 0x00, 0xD8, 0x0A]
                                power_on = [0x01, 0x06, 0x00, 0x01, 0x00, 0x01, 0x19, 0xCA]
                                vol_24v = [0x01,0x10,0x00,0x20,0x00,0x02,0x04,0x00,0x00,0x5D,0xC0,0xC9,0x77]
                                cur_2a = [0x01,0x10,0x00,0x22,0x00,0x02,0x04,0x00,0x00,0x4E,0x20,0x44,0x16]
                                self.powerControl( baudRate=9600,transmitData = power_off)
                                if not self.isCancelAllTest:
                                    self.result_signal.emit('设备已断电。等待3秒后自动重新上电。\n')
                                    for dd in range(3):
                                        self.result_signal.emit(f'剩余等待{3 - dd}秒……\n')
                                        time.sleep(1)
                                    self.powerControl(baudRate=9600, transmitData=vol_24v)
                                    if not self.isCancelAllTest:
                                        self.result_signal.emit('设置电压为24V。\n')
                                        self.powerControl(baudRate=9600, transmitData=cur_2a)
                                        if not self.isCancelAllTest:
                                            self.result_signal.emit('设置电压为2A。\n')
                                            self.powerControl(baudRate=9600, transmitData=power_on)
                                            if not self.isCancelAllTest:
                                                self.result_signal.emit('设备重新上电。\n')
                                                if not self.isCancelAllTest:
                                                    time.sleep(6)
                                                    # 数据检查
                                                    self.CPU_powerOffSaveTest(serial_transmitData=serial_transmitData1)
                                                    if not self.isCancelAllTest:
                                                        # 数据清除
                                                        self.CPU_powerOffSaveTest(serial_transmitData=serial_transmitData2)
                                # 手动开断电
                                # self.messageBox_signal.emit(['操作提示','请断开设备电源。'])
                                # reply= self.result_queue.get()
                                # if reply == QMessageBox.AcceptRole:
                                #     for t in range(3):
                                #         if t == 0:
                                #             self.result_signal.emit(f'等待3秒。\n')
                                #         self.result_signal.emit(f'剩余等待{3 - t}秒……\n')
                                #         time.sleep(1)
                                #     self.messageBox_signal.emit(['操作提示', '请接通设备电源。'])
                                #     reply = self.result_queue.get()
                                #     if reply == QMessageBox.AcceptRole:
                                #         for t in range(3):
                                #             if t == 0:
                                #                 self.result_signal.emit(f'等待3秒。\n')
                                #             self.result_signal.emit(f'剩余等待{3 - t}秒……\n')
                                #             time.sleep(1)
                                #         # 数据检查
                                #         self.CPU_powerOffSaveTest(serial_transmitData=serial_transmitData1)
                                #         if not self.isCancelAllTest:
                                #             # 数据清除
                                #             self.CPU_powerOffSaveTest(serial_transmitData=serial_transmitData2)
                                #     else:
                                #         self.messageBox_signal.emit(['测试警告', '未按提示接通设备电源，测试停止。'])
                                #         self.isPassPowerOffSave = False
                                #         self.cancelAllTest()
                                #
                                # else:
                                #     self.messageBox_signal.emit(['测试警告', '未按提示断开设备电源，测试停止。'])
                                #     self.isPassPowerOffSave = False
                                #     self.cancelAllTest()

                        except:
                            self.isPassPowerOffSave = False
                            self.showErrorInf('掉电保存测试')
                            self.cancelAllTest()
                        finally:
                            self.CPU_passItem[i] = self.isPassPowerOffSave
                            self.CPU_testItem[i] = True
                            #若工装的电是通过CPU供电的，断电后需要重新分配节点
                            self.configCANAddr(self.CANAddr1, self.CANAddr2, self.CANAddr3)
                            self.isPassAll &= self.isPassPowerOffSave
                            self.testNum = self.testNum - 1
                            if self.isPassPowerOffSave:
                                self.result_signal.emit("掉电保存测试通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("掉电保存测试未通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break


                    elif i == 7:#RTC
                        self.moveToRow_signal.emit([i, 0])
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------RTC测试----------------')
                        try:
                            if not self.CPU_isTest[6]:
                                testStartTime = time.time()
                                self.CPU_RTCTest('write')
                                if not self.isCancelAllTest:
                                    self.messageBox_signal.emit(['操作提示', '请断开设备电源。'])
                                    reply = self.result_queue.get()
                                    if reply == QMessageBox.AcceptRole:
                                        for t in range(3):
                                            if t == 0:
                                                self.result_signal.emit(f'等待3秒。\n')
                                            self.result_signal.emit(f'剩余等待{3 - t}秒……\n')
                                            time.sleep(1)
                                        self.messageBox_signal.emit(['操作提示', '请接通设备电源。'])
                                        reply = self.result_queue.get()
                                        if reply == QMessageBox.AcceptRole:
                                            for t in range(3):
                                                if t == 0:
                                                    self.result_signal.emit(f'等待3秒。\n')
                                                self.result_signal.emit(f'剩余等待{3 - t}秒……\n')
                                                time.sleep(1)

                                            self.CPU_RTCTest('read')
                                            if not self.isCancelAllTest:
                                                self.CPU_RTCTest('battery')
                                else:
                                    self.isPassRTC &= False
                            else:
                                self.CPU_RTCTest('read')
                                if not self.isCancelAllTest:
                                    self.CPU_RTCTest('battery')
                        except:
                            self.isPassRTC = False
                            self.showErrorInf('RTC测试')
                            self.cancelAllTest()
                        finally:
                            self.CPU_passItem[i] = self.isPassRTC
                            self.CPU_testItem[i] = True
                            if not self.CPU_isTest[6]:
                                # 若工装的电是通过CPU供电的，断电后需要重新分配节点
                                self.configCANAddr(self.CANAddr1, self.CANAddr2, self.CANAddr3)
                            self.isPassAll &= self.isPassRTC
                            self.testNum = self.testNum - 1
                            if self.isPassRTC:
                                self.result_signal.emit("RTC测试通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("RTC测试未通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 8:#FPGA
                        self.moveToRow_signal.emit([i, 0])
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------FPGA测试----------------')
                        try:
                            self.messageBox_signal.emit(["操作提示","请插入U盘，并观察HOST指示灯是否点亮？"])
                            reply = self.result_queue.get()
                            if reply == QMessageBox.AcceptRole:
                                transmitData0 = [0xAC, 6, 0x00, 0x04, 0x53, 0x00,
                                                 self.getCheckNum([0xAC, 6, 0x00, 0x04, 0x53, 0x00])]
                                transmitData1 = [0xAC, 6, 0x00, 0x04, 0x53, 0x01,
                                                 self.getCheckNum([0xAC, 6, 0x00, 0x04, 0x53, 0x01])]
                                transmitData2 = [0xAC, 6, 0x00, 0x04, 0x53, 0x02,
                                                 self.getCheckNum([0xAC, 6, 0x00, 0x04, 0x53, 0x02])]
                                transmitData_array = [transmitData0,transmitData1,transmitData2]
                                self.CPU_FPGATest(transmitData=transmitData0)
                                if not self.isCancelAllTest:
                                    if self.isPassFPGA:
                                        self.result_signal.emit('<保存固件>成功。\n\n')
                                    else:
                                        self.result_signal.emit('<保存固件>失败。\n\n')
                                    self.CPU_FPGATest(transmitData=transmitData1)
                                    if not self.isCancelAllTest:
                                        if self.isPassFPGA:
                                            self.result_signal.emit('<加载固件>成功。\n\n')
                                        else:
                                            self.result_signal.emit('<加载固件>失败。\n\n')
                                        self.CPU_FPGATest(transmitData=transmitData2)
                                        if not self.isCancelAllTest:
                                            if self.isPassFPGA:
                                                self.result_signal.emit('<端口配置>成功。\n\n')
                                            else:
                                                self.result_signal.emit('<端口配置>失败。\n\n')
                                #         else:
                                #             break
                                #     else:
                                #         break
                                # else:
                                #     break
                            else:
                                self.messageBox_signal.emit(['测试警告', '无法识别U盘，FPGA无法测试，是否取消后续测试？'])
                                self.isPassFPGA = False
                                reply = self.result_queue.get()
                                if reply == QMessageBox.AcceptRole:
                                    self.cancelAllTest()

                        except:
                            self.isPassFPGA = False
                            self.showErrorInf('FPGA')
                            self.cancelAllTest()
                        finally:
                            self.CPU_passItem[i] = self.isPassFPGA
                            self.CPU_testItem[i] = True
                            self.isPassAll &= self.isPassFPGA
                            self.testNum = self.testNum - 1
                            if self.isPassFPGA:
                                self.result_signal.emit("FPGA测试通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("FPGA测试未通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 9:#各指示灯
                        self.stopChannel = 2
                        self.moveToRow_signal.emit([i, 0])
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------各指示灯测试----------------')
                        try:
                            w_transmitData = [0xAC, 7, 0x00, 0x09, 0x10, 0x00, 0x01,
                                              self.getCheckNum([0xAC, 7, 0x00, 0x09, 0x10, 0x00, 0x01])]
                            r_transmitData = [0xAC, 6, 0x00, 0x09, 0x0E, 0x00,
                                              self.getCheckNum([0xAC, 6, 0x00, 0x09, 0x0E, 0x00])]
                            inTest_transmitData = [0xAC, 6, 0x00, 0x09, 0x10, 0x53,0xe1]
                                                   # self.getCheckNum([0xAC, 6, 0x00, 0x09, 0x10, 0x53])]
                            outTest_transmitData = [0xAC, 6, 0x00, 0x09, 0x10, 0x50,0xe4]
                                                    # self.getCheckNum([0xAC, 6, 0x00, 0x09, 0x10, 0x50])]
                            name_LED = ['RUN灯','ERROR灯','BAT_LOW灯','PRG灯','RS-232C灯','RS-485灯','HOST灯','所有灯']
                            #进入灯测试模式
                            self.CPU_LEDTest(transmitData=inTest_transmitData,LED_name='')
                            if not self.isCancelAllTest:
                                LED_thread = threading.Thread(target=self.CPU_LED_light_loop)
                                LED_thread.start()
                                image_CPU_LED = self.CPU_pic_dir + '/CPU_LED.gif'
                                self.gif_messageBox_signal.emit(
                                    [f'LED检测', f'LED指示灯是否如图所示循环点亮？', image_CPU_LED])
                                reply = self.result_queue.get()
                                if reply == QMessageBox.AcceptRole:
                                    self.isPassLED &= True
                                else:
                                    self.isPassLED &= False
                                    self.checkBox_messageBox_signal.emit(['故障LED灯选择', '请选择存在故障的LED灯。', 'CPU_LED'])

                                    messageList = self.result_queue.get()
                                    reply = messageList[0]
                                    CPU_LED_status = messageList[1]
                                    if reply == QMessageBox.AcceptRole:
                                        for cls in range(len(CPU_LED_status)):
                                            if not CPU_LED_status[cls]:
                                                self.isPassLED &= False
                                                self.result_signal.emit(f'{CPU_LED_status[cls]}存在问题。\n\n')

                                # 等待子线程执行结束
                                LED_thread.join()
                                w_allLED = [0xAC, 8, 0x00, 0x09, 0x10, 0xAA, 0x00, 0x00, 0x88]
                                self.CPU_LEDTest(transmitData=w_allLED, LED_name='所有灯')
                                # for l in range(8):
                                #     if l != 7:
                                #         self.result_signal.emit(f"----------开始测试{name_LED[l]}----------\n")
                                #         w_transmitData[5] = l
                                #         w_transmitData[7] = self.getCheckNum(w_transmitData[:7])
                                #         self.CPU_LEDTest(transmitData=w_transmitData, LED_name=name_LED[l])
                                #     else:
                                #         #关闭所有LED
                                #         w_allLED = [0xAC, 8, 0x00, 0x09, 0x10, 0xAA, 0x00, 0x00, 0x88]
                                #         self.CPU_LEDTest(transmitData=w_allLED, LED_name=name_LED[l])
                                #     if not self.isCancelAllTest:
                                #         if l != 7:
                                #             r_transmitData[5] = l
                                #             r_transmitData[6] = self.getCheckNum(r_transmitData[:6])
                                #             self.CPU_LEDTest(transmitData=r_transmitData, LED_name=name_LED[l])
                                #         else:
                                #             r_allLED = [0xAC, 6, 0x00, 0x09, 0x0E, 0xAA, 0x8c]
                                #             self.CPU_LEDTest(transmitData=r_allLED, LED_name=name_LED[l])
                                #
                                #     else:
                                #         break
                        except:
                            self.isPassLED =False
                            self.showErrorInf('LED测试')
                            self.cancelAllTest()
                        finally:
                            self.CPU_passItem[i] = self.isPassLED
                            self.CPU_testItem[i] = True
                            self.testNum = self.testNum - 1
                            # 退出灯测试模式
                            self.CPU_LEDTest(transmitData=outTest_transmitData, LED_name='')
                            if self.isPassLED:
                                self.result_signal.emit("LED测试通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("LED测试未通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 10:#本体IN
                        self.moveToRow_signal.emit([i, 0])
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------本体输入测试----------------')
                        r_transmitData = [0xAC, 7,0x00, 0x0A, 0x0E, 0x00, 0xAA,
                                          self.getCheckNum([0xAC, 7,0x00, 0x0A, 0x0E, 0x00, 0xAA])]
                        if not self.CPU_isTest[8]:
                            transmitData1 = [0xAC, 6, 0x00, 0x04, 0x53, 0x01,
                                             self.getCheckNum([0xAC, 6, 0x00, 0x04, 0x53, 0x01])]
                            transmitData2 = [0xAC, 6, 0x00, 0x04, 0x53, 0x02,
                                             self.getCheckNum([0xAC, 6, 0x00, 0x04, 0x53, 0x02])]
                            self.CPU_FPGATest(transmitData=transmitData1)
                            if self.isPassFPGA:
                                self.result_signal.emit('<保存固件>成功。\n\n')
                            else:
                                self.result_signal.emit('<保存固件>失败。\n\n')
                            if not self.isCancelAllTest:
                                self.CPU_FPGATest(transmitData=transmitData2)
                                if not self.isCancelAllTest:
                                    if self.isPassFPGA:
                                        self.result_signal.emit('<端口配置>成功。\n\n')
                                    else:
                                        self.result_signal.emit('<端口配置>失败。\n\n')
                        try:
                            CAN_init([1])
                            #先控制两块QN0016模块向CPU输入端口发送高电平
                            self.m_transmitData = [0xAA,0xAA,0x00,0x00,0x00,0x00,0x00,0x00]
                            CAN_option.transmitCAN(0x200 + self.CANAddr2, self.m_transmitData,1)
                            time.sleep(0.1)
                            CAN_option.transmitCAN(0x200 + self.CANAddr3, self.m_transmitData,1)
                            #1.读输入端口
                            self.CPU_InTest(transmitData=r_transmitData)
                            if not self.isCancelAllTest:
                                self.m_transmitData = [0x55, 0x55, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
                                CAN_option.transmitCAN(0x200 + self.CANAddr2, self.m_transmitData, 1)
                                time.sleep(0.1)
                                CAN_option.transmitCAN(0x200 + self.CANAddr3, self.m_transmitData, 1)
                                # 1.读输入端口
                                self.CPU_InTest(transmitData=r_transmitData)
                                if not self.isCancelAllTest:
                                    self.m_transmitData = [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
                                    CAN_option.transmitCAN(0x200 + self.CANAddr2, self.m_transmitData, 1)
                                    time.sleep(0.1)
                                    CAN_option.transmitCAN(0x200 + self.CANAddr3, self.m_transmitData, 1)
                                    # 2.读输入端口
                                    self.CPU_InTest(transmitData=r_transmitData)
                        except:
                            self.isPassIn = False
                            self.showErrorInf('本体输入测试')
                            self.cancelAllTest()
                        finally:
                            self.CPU_passItem[i] = self.isPassIn
                            self.CPU_testItem[i] = True
                            self.isPassAll &= self.isPassIn
                            self.testNum = self.testNum - 1
                            if self.isPassIn:
                                self.result_signal.emit("本体输入测试通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("本体输入测试未通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 11:#本体OUT
                        self.moveToRow_signal.emit([i, 0])
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------本体输出测试----------------')
                        w_transmitData = [0xAC, 10, 0x00, 0x0A, 0x10, 0x01, 0xAA, 2, 0xaa, 0xaa,
                              self.getCheckNum([0xAC, 10, 0x00, 0x0A, 0x10, 0x01, 0xAA, 2, 0xaa, 0xaa])]
                        if not self.CPU_isTest[8] and not self.CPU_isTest[10]:
                            transmitData1 = [0xAC, 6, 0x00, 0x04, 0x53, 0x01,
                                             self.getCheckNum([0xAC, 6, 0x00, 0x04, 0x53, 0x01])]
                            transmitData2 = [0xAC, 6, 0x00, 0x04, 0x53, 0x02,
                                             self.getCheckNum([0xAC, 6, 0x00, 0x04, 0x53, 0x02])]
                            self.CPU_FPGATest(transmitData=transmitData1)
                            if self.isPassFPGA:
                                self.result_signal.emit('<保存固件>成功。\n\n')
                            else:
                                self.result_signal.emit('<保存固件>失败。\n\n')
                            if not self.isCancelAllTest:
                                self.CPU_FPGATest(transmitData=transmitData2)
                                if not self.isCancelAllTest:
                                    if self.isPassFPGA:
                                        self.result_signal.emit('<端口配置>成功。\n\n')
                                    else:
                                        self.result_signal.emit('<端口配置>失败。\n\n')
                        try:
                            CAN_init([1])
                            self.CPU_OutTest(transmitData=w_transmitData)
                            if not self.isCancelAllTest:
                                w_transmitData = [0xAC, 10, 0x00, 0x0A, 0x10, 0x01, 0xAA, 2, 0x55, 0x55,
                                                  self.getCheckNum(
                                                      [0xAC, 10, 0x00, 0x0A, 0x10, 0x01, 0xAA, 2, 0x55, 0x55])]
                                self.CPU_OutTest(transmitData=w_transmitData)
                                if not self.isCancelAllTest:
                                    w_transmitData = [0xAC, 10, 0x00, 0x0A, 0x10, 0x01, 0xAA, 2, 0x00, 0x00,
                                      self.getCheckNum([0xAC, 10, 0x00, 0x0A, 0x10, 0x01, 0xAA, 2, 0x00, 0x00])]
                                    if not self.isCancelAllTest:
                                        self.CPU_OutTest(transmitData=w_transmitData)
                        except:
                            self.isPassOut = False
                            self.showErrorInf('本体输出测试')
                            self.cancelAllTest()
                        finally:
                            self.CPU_passItem[i] = self.isPassOut
                            self.CPU_testItem[i] = True
                            self.isPassAll &= self.isPassOut
                            self.testNum = self.testNum - 1
                            if self.isPassOut:
                                self.result_signal.emit("本体输出测试通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("本体输出测试未通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 12:#以太网
                        self.moveToRow_signal.emit([i, 0])
                        self.item_signal.emit([i, 1, 0, ''])
                        testStartTime = time.time()
                        self.result_signal.emit(f'----------------以太网测试----------------')
                        # defaultIP_0 = 192
                        # defaultIP_1 = 168
                        # defaultIP_2 = 1
                        # defaultIP_3 = 66
                        # defaultUDP_LSB = 1213
                        # defaultUDP_MSB = 50000
                        #
                        # r_transmitData_IP = [0xAC, 9, 0x00, 0x0B, 0x0E, defaultIP_0,
                        #                   defaultIP_1, defaultIP_2,defaultIP_3,
                        #                      self.getCheckNum([0xAC, 9, 0x00, 0x0B, 0x0E,
                        #                    defaultIP_0,defaultIP_1, defaultIP_2,defaultIP_3])]
                        # r_transmitData_UDP = [0xAC, 7, 0x00, 0x0B, 0x0E, defaultUDP_LSB,
                        #                       defaultUDP_MSB]
                        try:
                            self.messageBox_signal.emit(['操作提示','请将网线插入任意一个网口'])
                            reply = self.result_queue.get()
                            if reply == QMessageBox.AcceptRole:
                                self.TCP_UDPTest()
                                if not self.isCancelAllTest:
                                    self.messageBox_signal.emit(['操作提示', '请将网线插入另一个网口'])
                                    reply = self.result_queue.get()
                                    if reply == QMessageBox.AcceptRole:
                                        self.TCP_UDPTest()
                                    else:
                                        self.result_signal.emit('网线未插入，测试不通过!')
                                        self.isPassETH &= False
                                else:
                                    pass
                            else:
                                self.result_signal.emit('网线未插入，测试不通过!')
                                self.isPassETH &= False

                        except:
                            self.isPassETH=False
                            self.showErrorInf('本体以太网测试')
                            self.cancelAllTest()
                        finally:
                            self.CPU_passItem[i] = self.isPassETH
                            self.CPU_testItem[i] = True
                            self.isPassAll &= self.isPassETH
                            self.testNum = self.testNum - 1
                            if self.isPassETH:
                                self.result_signal.emit('本体以太网测试通过。'+ self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit('本体以太网测试未通过。'+ self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 13:#RS-232C
                        self.moveToRow_signal.emit([i, 0])
                        self.item_signal.emit([i, 1, 0, ''])
                        testStartTime = time.time()
                        self.result_signal.emit(f'----------------本体232测试----------------')
                        # 115200 57600 38400 19200 9600  4800
                        byte0_array = [0x00, 0x00, 0x00, 0x00, 0x80, 0xC0]
                        byte1_array = [0xC2, 0xE1, 0x96, 0x4B, 0x25, 0x12]
                        byte2_array = [0x01, 0x00, 0x00, 0x00, 0x00, 0x00]
                        byte3_array = [0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
                        w_transmitData = [0xAC, 10, 0x00, 0x0C, 0x10, 0x00,
                                             0x00, 0x00, 0x00, 0x00, 0x00]
                        r_transmitData = [0xAC, 6, 0x00, 0x0C, 0x0E, 0x00,
                                          self.getCheckNum([0xAC, 6, 0x00, 0x0C, 0x0E, 0x00])]
                        baudRate = [115200,57600,38400,19200,9600,4800]
                        try:
                            for b in range(6):
                                self.result_signal.emit(f'----------设置波特率为{baudRate[b]}----------\n')
                                w_transmitData[6] = byte0_array[b]
                                w_transmitData[7] = byte1_array[b]
                                w_transmitData[8] = byte2_array[b]
                                w_transmitData[9] = byte3_array[b]
                                w_transmitData[10] = self.getCheckNum(w_transmitData[:10])
                                self.CPU_232Test(transmitData=w_transmitData,
                                                 baudRate=[byte0_array[b],byte1_array[b],byte2_array[b],byte3_array[b]])
                                if not self.isCancelAllTest:
                                    self.CPU_232Test(transmitData=r_transmitData,
                                                 baudRate=[byte0_array[b],byte1_array[b],byte2_array[b],byte3_array[b]])
                                    if not self.isCancelAllTest:
                                        data= [0x00,0x11,0x22,0x33,0x44,0x55,0x55,0x66,0x77,0x88,0x99]
                                        if self.transmitBy232or485(type='232',baudRate=baudRate[b],
                                                                transmitData=data):
                                            self.isPass232 &= True
                                        else:
                                            self.isPass232 &= False
                                        if self.isCancelAllTest:
                                            break
                                    else:
                                        break
                                else:
                                    break
                        except:
                            self.isPass232=False
                            self.showErrorInf('本体232测试')
                            self.cancelAllTest()
                        finally:
                            self.CPU_passItem[i] = self.isPass232
                            self.CPU_testItem[i] = True
                            self.isPassAll &= self.isPass232
                            self.testNum = self.testNum - 1
                            if self.isPass232:
                                self.result_signal.emit('本体232测试通过。'+self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit('本体232测试未通过。'+self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 14:#RS-485
                        self.moveToRow_signal.emit([i, 0])
                        self.item_signal.emit([i, 1, 0, ''])
                        testStartTime = time.time()
                        self.result_signal.emit(f'----------------本体485测试----------------')
                        # 115200 57600 38400 19200 9600  4800
                        byte0_array = [0x00, 0x00, 0x00, 0x00, 0x80, 0xC0]
                        byte1_array = [0xC2, 0xE1, 0x96, 0x4B, 0x25, 0x12]
                        byte2_array = [0x01, 0x00, 0x00, 0x00, 0x00, 0x00]
                        byte3_array = [0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
                        w_transmitData = [0xAC, 10, 0x00, 0x0D, 0x10, 0x00,
                                          0x00, 0x00, 0x00, 0x00, 0x00]
                        r_transmitData = [0xAC, 6, 0x00, 0x0D, 0x0E, 0x00,
                                          self.getCheckNum([0xAC, 6, 0x00, 0x0D, 0x0E, 0x00])]
                        baudRate = [115200, 57600, 38400, 19200, 9600, 4800]
                        try:
                            for b in range(6):
                                self.result_signal.emit(f'----------设置波特率为{baudRate[b]}----------\n')
                                w_transmitData[6] = byte0_array[b]
                                w_transmitData[7] = byte1_array[b]
                                w_transmitData[8] = byte2_array[b]
                                w_transmitData[9] = byte3_array[b]
                                w_transmitData[10] = self.getCheckNum(w_transmitData[:10])
                                self.CPU_485Test(transmitData=w_transmitData,
                                                 baudRate=[byte0_array[b], byte1_array[b], byte2_array[b],
                                                           byte3_array[b]])
                                if not self.isCancelAllTest:
                                    self.CPU_485Test(transmitData=r_transmitData,
                                                     baudRate=[byte0_array[b], byte1_array[b], byte2_array[b],
                                                               byte3_array[b]])
                                    if not self.isCancelAllTest:
                                        data = [0x00, 0x11, 0x22, 0x33, 0x44, 0x55, 0x55, 0x66, 0x77, 0x88, 0x99]
                                        if self.transmitBy232or485(type='485', baudRate=baudRate[b],
                                                                   transmitData=data):
                                            self.isPass485 &= True
                                        else:
                                            self.isPass485 &= False
                                        if self.isCancelAllTest:
                                            break
                                    else:
                                        break
                                else:
                                    break
                        except:
                            self.isPass485=False
                            self.showErrorInf('本体485测试')
                            self.cancelAllTest()
                        finally:
                            self.CPU_passItem[i] = self.isPass485
                            self.CPU_testItem[i] = True
                            self.isPassAll &= self.isPass485
                            self.testNum = self.testNum - 1
                            if self.isPass485:
                                self.result_signal.emit('本体485测试通过。' + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit('本体485测试未通过。' + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 15:#右扩CAN
                        self.moveToRow_signal.emit([i, 0])
                        self.item_signal.emit([i, 1, 0, ''])
                        testStartTime = time.time()

                        transmitData_reset = [0xAC,7, 0x00, 0x0E, 0x10, 0x00, 0x00,
                                              self.getCheckNum([0xAC,7, 0x00, 0x0E, 0x10, 0x00, 0x00])]
                        transmitData_config = [0xAC, 7, 0x00, 0x0E, 0x10, 0x01, 0x01,
                                               self.getCheckNum([0xAC,7, 0x00, 0x0E, 0x10, 0x01, 0x01])]
                        #先把RESET灯和使能灯全熄灭
                        self.CPU_rightCANTest(transmitData=transmitData_reset,justOff=True)
                        self.CPU_rightCANTest(transmitData=transmitData_config,justOff=True)
                        try:
                            transmitData_reset[6] = 0x01
                            transmitData_reset[7] = self.getCheckNum(transmitData_reset[:7])
                            self.CPU_rightCANTest(transmitData=transmitData_reset)
                            if not self.isCancelAllTest:
                                transmitData_reset[6] = 0x00
                                transmitData_reset[7] = self.getCheckNum(transmitData_reset[:7])
                                self.CPU_rightCANTest(transmitData=transmitData_reset)
                                if not self.isCancelAllTest:
                                    transmitData_config[6] = 0x00
                                    transmitData_config[7] = self.getCheckNum(transmitData_config[:7])
                                    self.CPU_rightCANTest(transmitData=transmitData_config)
                                    if not self.isCancelAllTest:
                                        transmitData_config[6] = 0x01
                                        transmitData_config[7] = self.getCheckNum(transmitData_config[:7])
                                        self.CPU_rightCANTest(transmitData=transmitData_config)
                                        if not self.isCancelAllTest:
                                            try:
                                                # 连接CAN设备CANalyst-II，并初始化两个CAN通道
                                                CAN_bool,mess = self.CAN_init([0])
                                                if CAN_bool:
                                                    self.result_signal.emit('CAN设备的CAN1通道启动成功。\n')
                                                else:
                                                    self.result_signal.emit('CAN设备的CAN1通道启动失败，后续测试全部取消。\n')
                                                    self.cancelAllTest()
                                                    break
                                            except Exception as e:
                                                self.result_signal.emit(
                                                    f'CAN_init error:{e}\nCAN设备的CAN1通道启动失败，后续测试全部取消。\n')
                                                self.messageBox_signal.emit(['警告',
                                                                             f'CAN_init error:{e}\nCAN设备的CAN1通道启动失败，'
                                                                             f'后续测试全部取消。\n'])
                                                reply = self.result_queue.get()
                                                if reply == QMessageBox.AcceptRole:
                                                    self.cancelAllTest()
                                                else:
                                                    self.cancelAllTest()
                                                break
                                            for loopNum in range(32):
                                                rightCAN_transmitData = [(x+8*loopNum) for x in range(8)]
                                                isBool, whatEver = CAN_option.transmitCAN(0x67F,rightCAN_transmitData,0)
                                                if isBool:
                                                    self.pauseOption()
                                                    if not self.is_running:
                                                        self.cancelAllTest()
                                                        break
                                                    self.result_signal.emit(f'第{loopNum+1}次右扩CAN数据发送成功。\n'
                                                                            f'发送的数据为：{rightCAN_transmitData}\n')
                                                    recTime = time.time()
                                                    while True:
                                                        if (time.time() - recTime)*1000 > 2000:
                                                            break
                                                        m_can_obj = self.getCANObj(receiveID= 0x5FF)
                                                        bool_receive, m_can_obj = CAN_option.receiveCANbyID(0x5FF,
                                                                                                    self.waiting_time,0)
                                                        if bool_receive:
                                                            break
                                                    # QApplication.processEvents()
                                                    if bool_receive == 'stopReceive':
                                                        self.cancelAllTest()
                                                        self.isPassRightCAN = False
                                                        break
                                                    if bool_receive:
                                                        byte_data = bytes(m_can_obj.Data)
                                                        decimal_data = [int.from_bytes(byte_data[i:i+1], byteorder='big')
                                                                        for i in range(len(byte_data))]  # 将字节数组类型参数转换成十进制数组
                                                        # self.result_signal.emit(f'decimal_data:{decimal_data}')
                                                        self.result_signal.emit(f'decimal_data:{decimal_data}')
                                                        if len(decimal_data) == 0:
                                                            self.isPassRightCAN &= False
                                                            break
                                                        for ii in range(len(rightCAN_transmitData)):
                                                            if int(rightCAN_transmitData[ii]) != int(decimal_data[ii]):

                                                                self.result_signal.emit(f'右扩CAN数据接收失败。\n'
                                                                                f'接收的数据为：{decimal_data}\n\n')
                                                                self.isPassRightCAN &= False
                                                                break
                                                    else:
                                                        self.result_signal.emit(f'右扩CAN数据接收失败。\n\n')
                                                        self.isPassRightCAN &= False
                                                        break

                                                    self.result_signal.emit(f'右扩CAN数据接收成功。\n'
                                                                            f'接收的数据为：{rightCAN_transmitData}\n')
                                                    self.isPassRightCAN &= True
                                                    continue

                                                else:
                                                    self.result_signal.emit('右扩CAN数据发送失败。\n')
                                                    self.isPassRightCAN = False
                                                    break
                                            if not self.isPassRightCAN:
                                                self.messageBox_signal.emit(
                                                    ['测试警告', '右扩CAN测试不通过，是否取消后续测试？'])
                                                reply = self.result_queue.get()
                                                if reply == QMessageBox.AcceptRole:
                                                    self.cancelAllTest()

                            #             else:
                            #                 pass
                            #         else:
                            #             pass
                            #     else:
                            #         pass
                            # else:
                            #     pass
                        except:
                            self.isPassRightCAN=False
                            self.showErrorInf('本体右扩CAN测试')
                            self.cancelAllTest()
                        finally:
                            self.CPU_passItem[i] = self.isPassRightCAN
                            self.CPU_testItem[i] = True
                            self.isPassAll &= self.isPassRightCAN
                            self.testNum = self.testNum - 1
                            if self.isPassRightCAN:
                                self.result_signal.emit('本体右扩CAN测试通过。' + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit('本体右扩CAN测试未通过。' + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 16:#MAC/三码写入
                        self.moveToRow_signal.emit([i,0])
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        #写MAC
                        now_MAC = datetime.datetime.now()
                        self.MAC_array = [(now_MAC.year-2000), now_MAC.month,
                                 now_MAC.day, now_MAC.hour, now_MAC.minute,
                                 now_MAC.second]
                        self.result_signal.emit(f'时间：{self.MAC_array}\n')
                        w_MAC = [0xAC, 13, 0x00, 0x03, 0x10, 0x00,6,
                                 (now_MAC.year-2000), now_MAC.month,
                                 now_MAC.day, now_MAC.hour, now_MAC.minute,
                                 now_MAC.second,
                                 0x00]
                        # for mac in range(len(self.module_MAC)):
                        #     w_MAC[7+mac] = ord(self.module_MAC[mac])
                        w_MAC[13] = self.getCheckNum(w_MAC[:13])
                        #读MAC
                        r_MAC = [0xAC, 6, 0x00, 0x03, 0x0E, 0x00,
                                 self.getCheckNum([0xAC, 6, 0x00, 0x03, 0x0E, 0x00])]
                        # 写PN
                        w_PN = [0xAC, 31, 0x00, 0x03, 0x10, 0x01, 24,
                                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                0x00]
                        for pn in range(len(self.module_pn)):
                            w_PN[7+pn] = ord(self.module_pn[pn])
                        w_PN[31] = self.getCheckNum(w_PN[:31])
                        # 读PN
                        r_PN = [0xAC, 6, 0x00, 0x03, 0x0E, 0x01,
                                self.getCheckNum([0xAC, 6, 0x00, 0x03, 0x0E, 0x01])]

                        # 写SN
                        w_SN = [0xAC, 23, 0x00, 0x03, 0x10, 0x02, 16,
                                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                0x00]
                        for sn in range(len(self.module_sn)):
                            w_SN[7 + sn] = ord(self.module_sn[sn])
                        w_SN[23] = self.getCheckNum(w_SN[:23])
                        # 读SN
                        r_SN = [0xAC, 6, 0x00, 0x03, 0x0E, 0x02,
                                self.getCheckNum([0xAC, 6, 0x00, 0x03, 0x0E, 0x02])]

                        # 写REV
                        w_REV = [0xAC, 15, 0x00, 0x03, 0x10, 0x03, 8,
                                 ord(self.module_rev[0]), ord(self.module_rev[1]),0,0,0,0,0,0,
                                 self.getCheckNum([0xAC, 15, 0x00, 0x03, 0x10, 0x03, 8,
                                 ord(self.module_rev[0]), ord(self.module_rev[1])])]
                        # 读PN
                        r_REV = [0xAC, 6, 0x00, 0x03, 0x0E, 0x03,
                                 self.getCheckNum([0xAC, 6, 0x00, 0x03, 0x0E, 0x03])]

                        try:
                            #写MAC
                            self.CPU_configTest(transmitData=w_MAC)
                            if not self.isCancelAllTest:
                                # 读MAC
                                self.CPU_configTest(transmitData=r_MAC)
                                if not self.isCancelAllTest:
                                    # 写PN
                                    self.CPU_configTest(transmitData=w_PN)
                                    if not self.isCancelAllTest:
                                        # 读PN
                                        self.CPU_configTest(transmitData=r_PN)
                                        if not self.isCancelAllTest:
                                            #写SN
                                            self.CPU_configTest(transmitData=w_SN)
                                            if not self.isCancelAllTest:
                                                # 读SN
                                                self.CPU_configTest(transmitData=r_SN)
                                                if not self.isCancelAllTest:
                                                    # 写REV
                                                    self.CPU_configTest(transmitData=w_REV)
                                                    if not self.isCancelAllTest:
                                                        # 读REV
                                                        self.CPU_configTest(transmitData=r_REV)

                        except:
                            self.isPassConfig=False
                            self.showErrorInf('三码与MAC地址写入测试')
                            self.cancelAllTest()
                        finally:
                            self.CPU_passItem[i] = self.isPassConfig
                            self.CPU_testItem[i] = True
                            self.isPassAll &= self.isPassConfig
                            self.testNum = self.testNum - 1
                            if self.isPassConfig:
                                self.result_signal.emit('三码与MAC地址写入测试通过。' + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit('三码与MAC地址写入测试未通过。' + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 17:#MA0202
                        self.moveToRow_signal.emit([i, 0])
                        opType = ''
                        testStartTime = time.time()
                        self.pauseOption()
                        if not self.is_running:
                            self.cancelAllTest()
                            break
                        self.item_signal.emit([i, 1, 0, ''])
                        self.pauseOption()

                        if not self.is_running:
                            self.cancelAllTest()
                            break
                        self.result_signal.emit(f'----------------选项板测试----------------')
                        #波特率        115200 57600 38400 19200 9600  4800
                        byte0_array = [0x00, 0x00, 0x00, 0x00, 0x80, 0xC0]
                        byte1_array = [0xC2, 0xE1, 0x96, 0x4B, 0x25, 0x12]
                        byte2_array = [0x01, 0x00, 0x00, 0x00, 0x00, 0x00]
                        byte3_array = [0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
                        baudRate = [115200, 57600, 38400, 19200, 9600, 4800]
                        #读选项版类型
                        r_exType = [0xAC, 7, 0x00, 0x0F, 0x0E, 0xFF,0x00,
                                    0x30]

                        try:
                            opType = self.CPU_optionPanel(transmitData=r_exType)
                            if opType == 'MA0202':
                                if self.testType == 'CPU':
                                    self.isPassOp &= True
                                elif self.testType == 'MA0202':
                                    # 写MA0202量程和码值
                                    w_MA0202_1 = [0xAC, 11, 0x00, 0x0F, 0x10, 0x00,
                                                0x00, 0x00, 0x01, 0x00,0x00,
                                                0x00]
                                    # 写MA0202三码
                                    # w_MA0202_2 = [0xAC, 11, 0x00, 0x0F, 0x10, 0x00,
                                    #               0x00, 0x00, 0x01, 0x00, 0x00,
                                    #               0x00]
                                    # 读MA0202量程和码值
                                    r_MA0202_1 = [0xAC, 9, 0x00, 0x0F, 0x0E, 0x00,0x00,0x00,0x00,0x00]
                                    # 读MA0202三码
                                    r_MA0202_2 = [0xAC, 8, 0x00, 0x0F, 0x0E, 0x00,0x00,0xF1,0x00]

                                            #   0-10V     0-20mA
                                    AI_range=[0x2A,0x00,0x34,0x00]
                                            #   0-10V
                                    AO_range=[0xE8,0x03]
                                    #写AI量程0-10v
                                    w_MA0202_1[6] = 0x00
                                    w_MA0202_1[7] = 0x00
                                    w_MA0202_1[8] = 0x01
                                    w_MA0202_1[9] = 0x2A
                                    w_MA0202_1[10] = 0x00
                                    w_MA0202_1[11] = self.getCheckNum(w_MA0202_1[:11])
                                    self.pauseOption()
                                    if not self.is_running:
                                        self.cancelAllTest()
                                        break
                                    #1 写AI通道1量程0-10V
                                    self.CPU_optionPanel(transmitData=w_MA0202_1)
                                    # 2 写AI通道2量程0-10V
                                    w_MA0202_1[8] = 0x02
                                    w_MA0202_1[11] = self.getCheckNum(w_MA0202_1[:11])
                                    self.pauseOption()
                                    if not self.is_running:
                                        self.cancelAllTest()
                                        break
                                    self.CPU_optionPanel(transmitData=w_MA0202_1)
                                    if not self.isCancelAllTest:
                                        self.pauseOption()
                                        if not self.is_running:
                                            self.cancelAllTest()
                                            break
                                        self.result_signal.emit('------------写选项板AI通道1、2量程0-10V成功------------\n\n')
                                        # 3 读AI通道1量程0-10V
                                        r_MA0202_1[6] = 0x00
                                        r_MA0202_1[7] = 0x00
                                        r_MA0202_1[8] = 0x01
                                        r_MA0202_1[9] = self.getCheckNum(r_MA0202_1[:9])
                                        self.pauseOption()
                                        if not self.is_running:
                                            self.cancelAllTest()
                                            break
                                        self.CPU_optionPanel(transmitData=r_MA0202_1,param1=[0x2A,0x00])
                                        r_MA0202_1[8] = 0x02
                                        r_MA0202_1[9] = self.getCheckNum(r_MA0202_1[:9])
                                        self.pauseOption()
                                        if not self.is_running:
                                            self.cancelAllTest()
                                            break
                                        self.CPU_optionPanel(transmitData=r_MA0202_1,param1=[0x2A,0x00])
                                        if not self.isCancelAllTest:
                                            self.pauseOption()
                                            if not self.is_running:
                                                self.cancelAllTest()
                                                break
                                            self.result_signal.emit('------------读选项板AI通道1、2量程0-10V成功------------\n\n')
                                            # 4 写AI量程通道1量程0-20mA
                                            w_MA0202_1 = [0xAC, 11, 0x00, 0x0F, 0x10, 0x00,0x00, 0x00, 0x01, 0x34, 0x00,
                                                          0x00]
                                            w_MA0202_1[11] = self.getCheckNum(w_MA0202_1[:11])
                                            self.pauseOption()
                                            if not self.is_running:
                                                self.cancelAllTest()
                                                break
                                            self.CPU_optionPanel(transmitData=w_MA0202_1)
                                            # 5 写AI量程通道2量程0-20mA
                                            w_MA0202_1 = [0xAC, 11, 0x00, 0x0F, 0x10, 0x00, 0x00, 0x00, 0x02, 0x34, 0x00,
                                                          0x00]
                                            w_MA0202_1[11] = self.getCheckNum(w_MA0202_1[:11])
                                            self.pauseOption()
                                            if not self.is_running:
                                                self.cancelAllTest()
                                                break
                                            self.CPU_optionPanel(transmitData=w_MA0202_1)
                                            if not self.isCancelAllTest:
                                                self.pauseOption()
                                                if not self.is_running:
                                                    self.cancelAllTest()
                                                    break
                                                self.result_signal.emit('------------写选项板AI通道1、2量程0-20mA成功------------\n\n')
                                                # 6 读AI通道1量程0-20mA
                                                r_MA0202_1[6] = 0x00
                                                r_MA0202_1[7] = 0x00
                                                r_MA0202_1[8] = 0x01
                                                r_MA0202_1[9] = self.getCheckNum(r_MA0202_1[:9])
                                                self.pauseOption()
                                                if not self.is_running:
                                                    self.cancelAllTest()
                                                    break
                                                self.CPU_optionPanel(transmitData=r_MA0202_1,param1=[0x34,0x00])
                                                # 7 读AI通道2量程0-20mA
                                                r_MA0202_1[8] = 0x02
                                                r_MA0202_1[9] = self.getCheckNum(r_MA0202_1[:9])
                                                self.pauseOption()
                                                if not self.is_running:
                                                    self.cancelAllTest()
                                                    break
                                                self.CPU_optionPanel(transmitData=r_MA0202_1,param1=[0x34,0x00])
                                                if not self.isCancelAllTest:
                                                    if not self.is_running:
                                                        self.cancelAllTest()
                                                        break
                                                    self.result_signal.emit('------------读选项板AI通道1、2量程0-20mA成功------------\n\n')
                                                    try:
                                                        #8控制AQ0004模块发送电流
                                                        CAN_bool,mess = self.CAN_init([1])
                                                    except Exception as e:
                                                        self.isPassOp &= False
                                                        if not self.is_running:
                                                            self.cancelAllTest()
                                                            break
                                                        self.result_signal.emit(
                                                            f'CAN_init error:{e}\nCAN设备通道1启动失败，后续测试全部取消。\n')
                                                        self.messageBox_signal.emit(['警告',
                                                                                     f'CAN_init error:{e}\nCAN设备通道1启动失败，'
                                                                                     f'后续测试全部取消。\n'])
                                                        reply = self.result_queue.get()
                                                        if reply == QMessageBox.AcceptRole:
                                                            self.cancelAllTest()
                                                        else:
                                                            self.cancelAllTest()
                                                        break
                                                    if CAN_bool:
                                                        # 0v-10v的理论值
                                                        voltageTheory_0010 = [int(0x6c00*1.17),0x6c00, (int)(0x6c00 * 0.75),
                                                                              (int)(0x6c00 * 0.5), (int)(0x6c00 * 0.25),
                                                                              0]  # 0x6c00 =>27648
                                                        arrayVol_0010 = ["11.7V测试","10V测试", "7.5V测试", "5V测试",
                                                                         "2.5V测试","0V测试"]
                                                        # 0mA-20mA的理论值
                                                        currentTheory_0020 = [(int)(0x6c00*1.17),0x6c00, (int)(0x6c00 * 0.75),
                                                                              (int)(0x6c00 * 0.5), (int)(0x6c00 * 0.25),
                                                                              0]  # 0x6c00 =>27648
                                                        arrayCur_0020 = ["23.4mA", "20mA", "15mA",
                                                                         "10mA", "5mA","0mA"]
                                                        AQ_transmit = (c_ubyte * 8)(0x00,0x00,0x00,0x00,
                                                                                  0x00,0x00,0x00,0x00)
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
                                                                if bool_range1&bool_range2:
                                                                    if not self.is_running:
                                                                        self.cancelAllTest()
                                                                        break
                                                                    self.result_signal.emit(
                                                                        f"------------成功修改AQ0004通道1、2量程为0-20mA------------\n\n")
                                                                else:
                                                                    if not self.is_running:
                                                                        self.cancelAllTest()
                                                                        break
                                                                    self.result_signal.emit(
                                                                        f"------------修改AQ0004通道1、2量程为0-20mA失败，取消后续测试------------\n\n")
                                                                    self.cancelAllTest()
                                                                    self.isPassOp &= False
                                                                    break
                                                                time.sleep(0.5)
                                                            for ch in range(2):
                                                                #发送0-20mA的电流
                                                                AQ_transmit=[0x2b,0x11,0x64, ch+1,
                                                                             (currentTheory_0020[va] & 0xff),
                                                                             ((currentTheory_0020[va] >> 8)& 0xff),
                                                                             0x00,0x00]
                                                                for asd in range(10):
                                                                    CAN_option.transmitCAN((0x600 + self.CANAddr5),
                                                                                           AQ_transmit,1)
                                                                    # time.sleep(0.1)
                                                                #等待一秒，等待输出稳定
                                                                self.result_signal.emit('\n等待输出稳定……\n')
                                                                time.sleep(1)
                                                                if not self.is_running:
                                                                    self.cancelAllTest()
                                                                    break
                                                                self.result_signal.emit(f'\n\n------------AQ0004通道{ch+1}'
                                                                                        f'输出电流为：{arrayCur_0020[va]}'
                                                                                        f'------------\n')
                                                                r_MA0202_1 = [0xAC, 9, 0x00, 0x0F, 0x0E, 0x00, 0x00,
                                                                              0x01,ch+1,0x00]
                                                                r_MA0202_1[9] = self.getCheckNum(r_MA0202_1[:9])
                                                                rate = round(float(27648/4095),2)
                                                                self.pauseOption()
                                                                if not self.is_running:
                                                                    self.cancelAllTest()
                                                                    break
                                                                self.CPU_optionPanel(transmitData=r_MA0202_1,
                                                                                     param1=[(int(currentTheory_0020[va]/rate) & 0xff),
                                                                                            ((int(currentTheory_0020[va]/rate) >> 8)& 0xff)])
                                                            if self.isCancelAllTest:
                                                                break
                                                    else:
                                                        self.isPassOp &=False
                                                        self.messageBox_signal.emit(mess)
                                                        reply = self.result_queue.get()
                                                        if reply == QMessageBox.AcceptRole:
                                                            self.cancelAllTest()
                                                        else:
                                                            self.cancelAllTest()
                                    #                     break
                                    #             else:
                                    #                 break
                                    #         else:
                                    #             break
                                    #     else:
                                    #         break
                                    # else:
                                    #     break
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
                                            self.cancelAllTest()
                                            break
                                        # 1 写AO通道1量程0-10V
                                        self.CPU_optionPanel(transmitData=w_MA0202_1)
                                        # 2 写AI通道2量程0-10V
                                        w_MA0202_1[8] = 0x02
                                        w_MA0202_1[11] = self.getCheckNum(w_MA0202_1[:11])
                                        self.pauseOption()
                                        if not self.is_running:
                                            self.cancelAllTest()
                                            break
                                        self.CPU_optionPanel(transmitData=w_MA0202_1)
                                        if not self.isCancelAllTest:
                                            if not self.is_running:
                                                self.cancelAllTest()
                                                break
                                            self.result_signal.emit(
                                                '------------写选项板AO通道1、2量程0-10V成功------------\n\n')
                                            # 3 读AO通道1量程0-10V
                                            r_MA0202_1[6] = 0x01
                                            r_MA0202_1[7] = 0x00
                                            r_MA0202_1[8] = 0x01
                                            r_MA0202_1[9] = self.getCheckNum(r_MA0202_1[:9])
                                            self.pauseOption()
                                            if not self.is_running:
                                                self.cancelAllTest()
                                                break
                                            self.CPU_optionPanel(transmitData=r_MA0202_1, param1=[0xE8, 0x03])
                                            r_MA0202_1[8] = 0x02
                                            r_MA0202_1[9] = self.getCheckNum(r_MA0202_1[:9])
                                            self.pauseOption()
                                            if not self.is_running:
                                                self.cancelAllTest()
                                                break
                                            self.CPU_optionPanel(transmitData=r_MA0202_1, param1=[0xE8, 0x03])
                                            if not self.isCancelAllTest:
                                                self.pauseOption()
                                                if not self.is_running:
                                                    self.cancelAllTest()
                                                    break
                                                self.result_signal.emit(
                                                    '------------读选项板AO通道1、2量程0-10V成功------------\n\n')
                                                try:
                                                    # 8控制选项板AO通道发送电压
                                                    CAN_bool, mess = self.CAN_init([1])
                                                    if CAN_bool:
                                                        # 0v-10v的理论值
                                                        voltageTheory_0010 = [int(4095 * 1.17),4095,
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
                                                                        self.cancelAllTest()
                                                                        break
                                                                    self.result_signal.emit(
                                                                        f"------------成功修改AE0400通道1、2量程为0-10V------------\n\n")
                                                                else:
                                                                    if not self.is_running:
                                                                        self.cancelAllTest()
                                                                        break
                                                                    self.result_signal.emit(
                                                                        f"------------修改AE0400通道1、2量程为0-10V失败，取消后续测试------------\n\n")
                                                                    self.cancelAllTest()
                                                                    self.isPassOp &= False
                                                                    break

                                                            for ch in range(2):
                                                                # 发送0-10V的电压
                                                                w_MA0202_1 = [0xAC, 11, 0x00, 0x0F, 0x10, 0x00,
                                                                              0x01, 0x01, ch+1,
                                                                              (voltageTheory_0010[va] & 0xff),
                                                                              ((voltageTheory_0010[va] >> 8) & 0xff),0x00]
                                                                if va == 5:
                                                                    w_MA0202_1[9] = 0x00
                                                                    w_MA0202_1[10] = 0x00
                                                                w_MA0202_1[11] = self.getCheckNum(w_MA0202_1[:11])
                                                                self.pauseOption()
                                                                if not self.is_running:
                                                                    self.cancelAllTest()
                                                                    break
                                                                self.CPU_optionPanel(transmitData=w_MA0202_1)
                                                                if not self.isCancelAllTest:
                                                                    if not self.is_running:
                                                                        self.cancelAllTest()
                                                                        break
                                                                    self.result_signal.emit(
                                                                        f'\n\n------------选项板AO通道{ch+1}'
                                                                        f'输出电压为：{arrayVol_0010[va]}'
                                                                        f'------------')
                                                                    if not self.is_running:
                                                                        self.cancelAllTest()
                                                                        break
                                                                    self.result_signal.emit('等待接收的信号稳定……\n\n')
                                                                    tt0 = time.time()
                                                                    while True:
                                                                        rate = round(float(27648 / 4095), 2)
                                                                        if (time.time() - tt0)*1000 >10000:
                                                                            if not self.is_running:
                                                                                self.cancelAllTest()
                                                                                break
                                                                            self.result_signal.emit(
                                                                                '\n\n*********AE0400无法接收到稳定数据，'
                                                                                '取消后续测试！*********\n\n')
                                                                            self.isPassOp &= False
                                                                            # self.cancelAllTest()
                                                                            break
                                                                        bool_AE,AE_value = self.receiveAIData(CANAddr_AI= self.CANAddr4,channel=ch)

                                                                        if abs(voltageTheory_0010[va] - int(AE_value/rate)) < 40:
                                                                            break
                                                                    if self.isCancelAllTest:
                                                                        break

                                                                    tt1 = time.time()
                                                                    while True:
                                                                        if (time.time() - tt1)*1000 >2000:
                                                                            self.isPassOp &= False
                                                                            if not self.is_running:
                                                                                self.cancelAllTest()
                                                                                break
                                                                            self.result_signal.emit('\n\n*********AE0400未接收到数据，取消后续测试！*********\n\n')
                                                                            # self.cancelAllTest()
                                                                            break
                                                                        try:
                                                                            bool_AE,AE_value = self.receiveAIData(CANAddr_AI= self.CANAddr4,channel=ch)
                                                                        except:
                                                                            if not self.is_running:
                                                                                self.cancelAllTest()
                                                                                break
                                                                            self.result_signal.emit(
                                                                                f'ErrorInf:\n{traceback.format_exc()}')
                                                                        if bool_AE:
                                                                            if not self.is_running:
                                                                                self.cancelAllTest()
                                                                                break
                                                                            self.result_signal.emit(
                                                                                f'-----------模拟量输出通道{ch+1}写入的值为{voltageTheory_0010[va]}-----------。\n')
                                                                            if not self.is_running:
                                                                                self.cancelAllTest()
                                                                                break

                                                                            self.result_signal.emit(
                                                                                f'-----------AE0400接收值转换比例后的值为{int(AE_value/rate)}。-----------\n')
                                                                            try:
                                                                                if abs(voltageTheory_0010[va] - int(AE_value/rate)) < 40:
                                                                                    self.isPassOp &= True
                                                                                    break
                                                                                else:
                                                                                    self.isPassOp &= False
                                                                                    break
                                                                            except:
                                                                                if not self.is_running:
                                                                                    self.cancelAllTest()
                                                                                    break

                                                                                self.result_signal.emit(
                                                                                    f'ErrorInf:\n{traceback.format_exc()}')
                                                                                break
                                                                        else:
                                                                            continue
                                                            if self.isCancelAllTest:
                                                                break
                                                    else:
                                                        self.isPassOp &= False
                                                        self.messageBox_signal.emit(mess)
                                                        reply = self.result_queue.get()
                                                        if reply == QMessageBox.AcceptRole:
                                                            self.cancelAllTest()
                                                        else:
                                                            self.cancelAllTest()
                                                        break
                                                except Exception as e:
                                                    self.isPassOp &= False
                                                    if not self.is_running:
                                                        self.cancelAllTest()
                                                        break

                                                    self.result_signal.emit(
                                                        f'CAN_init error:{e}\nCAN设备通道1启动失败，后续测试全部取消。\n')
                                                    self.messageBox_signal.emit(['警告',
                                                                                 f'CAN_init error:{e}\nCAN设备通道1启动失败，'
                                                                                 f'后续测试全部取消。\n'])

                                                    reply = self.result_queue.get()
                                                    if reply == QMessageBox.AcceptRole:
                                                        self.cancelAllTest()
                                                    else:
                                                        self.cancelAllTest()
                                                    break
                                    else:
                                        break

                                    # 三码
                                    if not self.isCancelAllTest:
                                        option_pn='P0000010390631'
                                        option_sn = 'S1223-001083'
                                        option_rev = '06'
                                        w_PN = [0xac, 0x21, 0x00, 0x0f, 0x10, 0x00, 0x00, 0xf1, 0x18,
                                                0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
                                                0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
                                                0x00]
                                        for pn in range(len(option_pn)):
                                            w_PN[9+pn] = ord(option_pn[pn])
                                        w_PN[33] = self.getCheckNum(w_PN[:33])
                                        if not self.is_running:
                                            self.cancelAllTest()
                                            break

                                        self.result_signal.emit(f'开始写入选项板PN码：{option_pn}。\n')
                                        self.pauseOption()
                                        if not self.is_running:
                                            self.cancelAllTest()
                                            break
                                        self.CPU_optionPanel(transmitData=w_PN)
                                        if not self.isCancelAllTest:
                                            if not self.is_running:
                                                self.cancelAllTest()
                                                break

                                            self.result_signal.emit('写入PN成功！\n')
                                            self.result_signal.emit('开始读取选项板PN码。\n\n')
                                            r_PN = [0xac, 8, 0x00, 0x0f, 0x0e, 0x00, 0x00, 0xf1, 0x00]
                                            r_PN[8] = self.getCheckNum(r_PN[:8])
                                            self.pauseOption()
                                            if not self.is_running:
                                                self.cancelAllTest()
                                                break
                                            self.CPU_optionPanel(transmitData=r_PN,param2=option_pn)
                                            if not self.isCancelAllTest:
                                                if not self.is_running:
                                                    self.cancelAllTest()
                                                    break

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
                                                    self.cancelAllTest()
                                                    break

                                                self.result_signal.emit(f'开始写入选项板SN码：{option_sn}。\n')
                                                self.pauseOption()
                                                if not self.is_running:
                                                    self.cancelAllTest()
                                                    break
                                                self.CPU_optionPanel(transmitData=w_SN)
                                                if not self.isCancelAllTest:
                                                    if not self.is_running:
                                                        self.cancelAllTest()
                                                        break

                                                    self.result_signal.emit('写入SN成功！\n')
                                                    self.result_signal.emit('开始读取选项板SN码。\n\n')
                                                    r_SN = [0xac, 8, 0x00, 0x0f, 0x0e, 0x00, 0x00, 0xf2, 0x00]
                                                    r_SN[8] = self.getCheckNum(r_SN[:8])
                                                    self.pauseOption()
                                                    if not self.is_running:
                                                        self.cancelAllTest()
                                                        break
                                                    self.CPU_optionPanel(transmitData=r_SN, param2=option_sn)
                                                    if not self.isCancelAllTest:
                                                        if not self.is_running:
                                                            self.cancelAllTest()
                                                            break

                                                        self.result_signal.emit('读SN成功！\n\n')
                                                        w_REV = [0xac, 0x11, 0x00, 0x0f, 0x10, 0x00, 0x00, 0xf3, 0x08,
                                                                 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                                                 0x00]
                                                        for rev in range(len(option_rev)):
                                                            w_REV[9 + rev] = ord(option_rev[rev])
                                                        w_REV[17] = self.getCheckNum(w_REV[:17])
                                                        # self.result_signal.emit(f'w_rev:{w_REV}')
                                                        if not self.is_running:
                                                            self.cancelAllTest()
                                                            break

                                                        self.result_signal.emit(f'开始写入选项板REV码：{option_rev}。\n')
                                                        self.pauseOption()
                                                        if not self.is_running:
                                                            self.cancelAllTest()
                                                            break
                                                        self.CPU_optionPanel(transmitData=w_REV)
                                                        if not self.isCancelAllTest:
                                                            if not self.is_running:
                                                                self.cancelAllTest()
                                                                break

                                                            self.result_signal.emit('写入REV成功！\n')
                                                            self.result_signal.emit('开始读取选项板REV码。\n\n')
                                                            r_REV = [0xac, 8, 0x00, 0x0f, 0x0e, 0x00, 0x00, 0xf3, 0x3b]
                                                            # r_REV[8] = self.getCheckNum(r_REV[:8])
                                                            self.pauseOption()
                                                            if not self.is_running:
                                                                self.cancelAllTest()
                                                                break
                                                            self.CPU_optionPanel(transmitData=r_REV, param2=option_rev)
                                                            if not self.isCancelAllTest:
                                                                if not self.is_running:
                                                                    self.cancelAllTest()
                                                                    break

                                                                self.result_signal.emit('读REV成功！\n\n')
                                                            else:
                                                                break
                                                        else:
                                                            break
                                                    else:
                                                        break
                                                else:
                                                    break
                                            else:
                                                break
                                        else:
                                            break
                                    else:
                                        break


                            elif opType == '422':
                                pass
                            elif opType == '485':
                                w_485 = [0xAC, 11, 0x00, 0x0F, 0x10, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
                                         0x00]
                                # 读MA0202
                                r_485 = [0xAC, 7, 0x00, 0x0F, 0x0E, 0x02, 0x00,
                                         0x2D]
                                for br in range(6):
                                    if not self.is_running:
                                        self.cancelAllTest()
                                        break

                                    self.result_signal.emit(f'----------设置波特率为{baudRate[br]}----------\n')
                                    w_485[7] = byte0_array[br]
                                    w_485[8] = byte1_array[br]
                                    w_485[9] = byte2_array[br]
                                    w_485[10] = byte3_array[br]
                                    w_485[11] = self.getCheckNum(w_485[:11])
                                    self.pauseOption()
                                    if not self.is_running:
                                        self.cancelAllTest()
                                        break
                                    self.CPU_optionPanel(transmitData=w_485,
                                                         baudRate=[byte0_array[br], byte1_array[br],
                                                                   byte2_array[br], byte3_array[br]])
                                    if not self.isCancelAllTest:
                                        self.pauseOption()
                                        if not self.is_running:
                                            self.cancelAllTest()
                                            break
                                        self.CPU_optionPanel(transmitData=r_485,
                                                             baudRate=[byte0_array[br], byte1_array[br],
                                                                       byte2_array[br], byte3_array[br]])
                                        if not self.isCancelAllTest:
                                            data = [i for i in range(200)]
                                            self.pauseOption()
                                            if not self.is_running:
                                                self.cancelAllTest()
                                                break
                                            if self.transmitBy232or485(type='485', baudRate=baudRate[br],
                                                                       transmitData=data):
                                                self.isPassOp &= True
                                            else:
                                                self.isPassOp &= False
                                            if self.isCancelAllTest:
                                                break
                                        else:
                                            break
                                    else:
                                        break
                            elif opType == '232':
                                w_232 = [0xAC, 11, 0x00, 0x0F, 0x10, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00,
                                         0x00]
                                # 读MA0202
                                r_232 = [0xAC, 7, 0x00, 0x0F, 0x0E, 0x03, 0x00,
                                         0x2C]
                                for br in range(6):
                                    if not self.is_running:
                                        self.cancelAllTest()
                                        break

                                    self.result_signal.emit(f'----------设置波特率为{baudRate[br]}----------\n')
                                    w_232[7] = byte0_array[br]
                                    w_232[8] = byte1_array[br]
                                    w_232[9] = byte2_array[br]
                                    w_232[10] = byte3_array[br]
                                    w_232[11] = self.getCheckNum(w_232[:11])
                                    self.pauseOption()
                                    if not self.is_running:
                                        self.cancelAllTest()
                                        break
                                    self.CPU_optionPanel(transmitData=w_232,
                                                         baudRate=[byte0_array[br],byte1_array[br],
                                                                   byte2_array[br],byte3_array[br]])
                                    if not self.isCancelAllTest:
                                        self.pauseOption()
                                        if not self.is_running:
                                            self.cancelAllTest()
                                            break
                                        self.CPU_optionPanel(transmitData=r_232,
                                                             baudRate=[byte0_array[br], byte1_array[br],
                                                                       byte2_array[br], byte3_array[br]])
                                        if not self.isCancelAllTest:
                                            data = [i for i in range(200)]
                                            self.pauseOption()
                                            if not self.is_running:
                                                self.cancelAllTest()
                                                break
                                            if self.transmitBy232or485(type='232', baudRate=baudRate[br],
                                                                       transmitData=data):
                                                self.isPassOp &= True
                                            else:
                                                self.isPassOp &= False
                                            if self.isCancelAllTest:
                                                break
                                        else:
                                            break
                                    else:
                                        break
                            else:
                                self.isPassOp &= False
                        except:
                            self.isPassOp=False
                            self.showErrorInf('选项板测试')
                            self.cancelAllTest()
                        finally:
                            self.CPU_passItem[i] = self.isPassOp
                            self.CPU_testItem[i] = True
                            if not self.is_running:
                                self.isPassOp &= False
                            self.isPassAll &= self.isPassOp
                            self.testNum = self.testNum - 1
                            if self.isPassOp:
                                self.result_signal.emit('选项板测试通过。' + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit('选项板测试未通过。' + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break

                if self.isCancelAllTest:#对应前面的for循环
                    if not self.is_running:
                        self.cancelAllTest()

                    self.result_signal.emit("后续测试已全部取消，测试结束。")
                    self.isPassAll = False
                    break
            if self.is_running:

                if self.isPassAll and self.testNum == 0:
                    self.messageBox_signal.emit(['操作提示', f'产品全部测试项合格,请烧录正式固件。\n'])
                    self.label_signal.emit(['pass', '全部通过'])
                elif self.isPassAll and self.testNum != 0:
                    self.messageBox_signal.emit(['操作提示', f'产品部分测试项合格,请决定是否烧录正式固件。\n'])
                    self.label_signal.emit(['testing', '部分通过'])
                elif not self.isPassAll:
                    self.messageBox_signal.emit(['操作提示', f'产品存在部分测试项不合格，请检查！\n'])
                    self.label_signal.emit(['fail', '未通过'])
                reply = self.result_queue.get()
                if reply == QMessageBox.AcceptRole or reply == QMessageBox.RejectRole:
                    self.messageBox_signal.emit(['操作提示', f'请将拨杆拨到STOP位置（下）。\n'])
                    reply = self.result_queue.get()
                    if reply == QMessageBox.AcceptRole or reply == QMessageBox.RejectRole:
                        pass
            else:
                self.messageBox_signal.emit(['操作提示', f'测试被中止！\n'])
                self.label_signal.emit(['fail', '测试停止'])

        except:
            self.isPassAll &= False
            self.result_signal.emit(f'测试出现问题，请检查测试程序和测试设备！\n'
                                                     f'"ErrorInf:\n{traceback.format_exc()}"')
            self.showErrorInf('测试')
            return False
        finally:
            self.result_signal.emit(f'self.testNum={self.testNum}\n')
            if True:
                self.result_signal.emit('开始生成校准校验表…………' + self.HORIZONTAL_LINE)
                code_array = [self.module_pn, self.module_sn, self.module_rev,self.module_MAC]
                station_array = [False,False,False]

                excel_bool, book, sheet, self.CPU_row = otherOption.generateExcel(code_array,
                                                                                 station_array, 0, 'CPU')
                if not excel_bool:
                    self.result_signal.emit('校准校验表生成出错！请检查代码！' + self.HORIZONTAL_LINE)
                else:
                    try:
                        self.fillInCPUData(True, book, sheet)
                        self.result_signal.emit('生成校准校验表成功' + self.HORIZONTAL_LINE)
                    except:
                        self.result_signal.emit('校准校验表生成出错！请检查代码！' + self.HORIZONTAL_LINE)
                        self.result_signal.emit(f"校验表错误信息:\n{traceback.format_exc()}\n")

            self.allFinished_signal.emit()
            self.pass_signal.emit(False)
    #CPU外观检查
    def CPU_appearanceTest(self):
        appearanceStart_time = time.time()
        self.item_signal.emit([0, 1, 0, ''])
        if not self.is_running:
            self.cancelAllTest()
            return False
        self.messageBox_signal.emit(['外观检测', '请检查：\n（1）外壳字体是否清晰?\n（2）型号是否正确？\n（3）外壳是否完好？'])
        reply = self.result_queue.get()
        if reply == QMessageBox.AcceptRole:
            self.isPassAppearance = True
            appearanceEnd_time = time.time()
            appearanceTest_time = round(appearanceEnd_time - appearanceStart_time, 2)
            self.item_signal.emit([0, 2, 1, appearanceTest_time])
        else:
            self.isPassAppearance = False
            appearanceEnd_time = time.time()
            appearanceTest_time = round(appearanceEnd_time - appearanceStart_time, 2)
            self.item_signal.emit([0, 2, 2, appearanceTest_time])

        self.testNum = self.testNum - 1

    #CPU型号检查
    def CPU_typeTest(self,serial_transmitData:list):
        isThisPass = True
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
            self.isPassTypeTest= False
            isThisPass = False
            self.cancelAllTest()
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime)*1000>self.waiting_time:
                self.isPassTypeTest &= False
                isThisPass = False
                break
            # 发送数据
            typeC_serial.write(bytes(serial_transmitData))
            # 等待0.5s后接收数据
            time.sleep(1)
            serial_receiveData = [0 for x in range(40)]
            # 接收数组数据
            data = typeC_serial.read(40)
            serial_receiveData = [hex(i) for i in data]
            for i in range(len(serial_receiveData)):
                serial_receiveData[i] = int(serial_receiveData[i],16)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                self.result_signal.emit('未接收到正确数据，再次接收！\n')
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                self.isPassTypeTest &= False
                isThisPass = False
                break
            elif dataLen == 8:
                if trueData[6] == int(0x00):#接收成功:
                    if trueData[5] == int(0x00):#点数
                        if trueData[7] == 40:
                            self.result_signal.emit("点数正确\n")
                            isThisPass = True
                            self.isPassTypeTest &= True
                            break
                        else:
                            self.result_signal.emit("点数错误\n")
                            isThisPass = False
                            self.isPassTypeTest &= False
                            break
                    elif trueData[5] == int(0x01):#类型
                        if trueData[7] == int(0x00):#NPN
                            self.result_signal.emit("类型正确\n")
                            self.isPassTypeTest &= True
                            isThisPass = True
                            break
                        else:
                            self.result_signal.emit("类型错误\n")
                            self.isPassTypeTest &= False
                            isThisPass = False
                            break
                elif trueData[6] == 0x01:#接收失败
                    self.isPassTypeTest &= False
                    isThisPass = False
                    break
            else:
                continue
        if not isThisPass:
            self.messageBox_signal.emit(['测试警告', '设备型号测试不通过，后续测试中止！'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole or reply == QMessageBox.RejectRole:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    #SRAM测试
    def CPU_SRAMTest(self,serial_transmitData:list):

        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
            self.isPassSRAM = False
            self.cancelAllTest()
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassSRAM &= False
                break
            # 发送数据
            typeC_serial.write(bytes(serial_transmitData))
            # 等待0.5s后接收数据
            time.sleep(3)
            serial_receiveData = [0 for x in range(40)]
            # 接收数组数据
            data = typeC_serial.read(40)
            serial_receiveData = [hex(i) for i in data]
            for i in range(len(serial_receiveData)):
                serial_receiveData[i] = int(serial_receiveData[i],16)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                self.result_signal.emit('未接收到正确数据，再次接收！\n')
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                self.isPassSRAM &= False
                break
            elif dataLen == 6:
                if trueData[5] == int(0x00):  # 接收成功
                    self.isPassSRAM &= True
                    break
                elif trueData[5] == int(0x01):  # 接收失败
                    self.isPassSRAM &= False
                    break
            else:
                continue
        if not self.isPassSRAM:
            self.messageBox_signal.emit(['测试警告', 'SRAM测试不通过，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()

        # 关闭串口
        typeC_serial.close()

    #FLASH测试
    def CPU_FLASHTest(self,serial_transmitData:list):
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
            self.isPassFLASH = False
            self.cancelAllTest()
        loopStartTime = time.time()

        while True:
            if (time.time() - loopStartTime) * 1000 > 50*1000:
                self.isPassFLASH &= False
                break
            # 发送数据
            typeC_serial.write(bytes(serial_transmitData))
            # 等待0.5s后接收数据
            for t in range(15):
                if t ==0:
                    self.pauseOption()
                    if not self.is_running:
                        self.isPassFLASH = False
                        self.cancelAllTest()
                        break
                    self.result_signal.emit(f'等待设备自检查。\n')
                self.result_signal.emit(f'剩余等待{15-t}秒……\n')
                time.sleep(1)
            serial_receiveData = [0 for x in range(40)]
            # 接收数组数据
            data = typeC_serial.read(40)
            serial_receiveData = [hex(i) for i in data]
            for i in range(len(serial_receiveData)):
                serial_receiveData[i] = int(serial_receiveData[i],16)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                self.result_signal.emit('未接收到正确数据，再次接收！\n')
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                self.isPassFLASH &= False
                break
            elif dataLen == 6:
                if trueData[5] == int(0x00):  # 接收成功
                    self.isPassFLASH &= True
                    break
                elif trueData[5] == int(0x01):  # 接收失败
                    self.isPassFLASH &= False
                    break
            else:
                continue
        if not self.isPassFLASH:
            self.messageBox_signal.emit(['测试警告', 'FLASH测试不通过，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()


        # 关闭串口
        typeC_serial.close()

    # FPGA测试
    def CPU_FPGATest(self,transmitData:list):
        isThisPass = True
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            if not self.is_running:
                self.cancelAllTest()
                return False
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
            self.isPassFPGA = False
            isThisPass = False
            self.cancelAllTest()
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassFPGA &= False
                isThisPass = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            name_dict ={0x00:'保存固件',0x01:'加载固件',0x02:'端口配平'}
            if not self.is_running:
                self.cancelAllTest()
                return False
            self.result_signal.emit(f'等待<{name_dict[transmitData[5]]}>。\n\n')
            # 等待0.5s后接收数据
            time.sleep(5)
            serial_receiveData = [0 for x in range(40)]
            # 接收数组数据
            data = typeC_serial.read(40)
            serial_receiveData = [hex(i) for i in data]
            for i in range(len(serial_receiveData)):
                serial_receiveData[i] = int(serial_receiveData[i],16)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 7:
                if trueData[6] == int(0x00):  # 成功
                    self.isPassFPGA &= True
                    isThisPass = True
                    break
                elif trueData[6] == int(0x01):  # 失败
                    self.isPassFPGA &= False
                    isThisPass = False
                    break
            else:
                continue

        if not self.isPassFPGA:
            if not self.is_running:
                self.cancelAllTest()
                return False
            self.messageBox_signal.emit(['测试警告', 'FPGA测试不通过，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()

        # 关闭串口
        typeC_serial.close()

    # 拨杆测试
    def CPU_LeverTest(self,state:int,serial_transmitData:list):
        isThisPass = True
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
            self.isPassLever = False
            isThisPass = False
            self.cancelAllTest()
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassLever &= False
                isThisPass = False
                break
            # 发送数据
            typeC_serial.write(bytes(serial_transmitData))
            # 等待0.5s后接收数据
            time.sleep(1)
            serial_receiveData = [0 for x in range(40)]
            # 接收数组数据
            data = typeC_serial.read(40)
            serial_receiveData = [hex(i) for i in data]
            for i in range(len(serial_receiveData)):
                serial_receiveData[i] = int(serial_receiveData[i],16)
            trueData,dataLen,isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                self.result_signal.emit('未接收到正确数据，再次接收！\n')
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                self.isPassLever &= False
                isThisPass = False
                break
            elif dataLen == 8: #成功
                if trueData[7] == int(state):  # 测试通过
                    self.isPassLever &= True
                    isThisPass = True
                    break
                else:  # 测试不通过
                    self.isPassLever &= False
                    isThisPass = False
                    break
            elif dataLen == 7:  # 失败
                self.isPassLever &= False
                isThisPass = False
                break
            else:
                continue
        if not isThisPass:
            self.messageBox_signal.emit(['测试警告', '拨杆测试不通过，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()

        # 关闭串口
        typeC_serial.close()

    # MFK测试
    def CPU_MFKTest(self,state:int,serial_transmitData:list):
        isThisPass = True
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
            self.isPassMFK = False
            isThisPass = False
            self.cancelAllTest()
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassMFK &= False
                isThisPass = False
                break
            # 发送数据
            typeC_serial.write(bytes(serial_transmitData))
            # 等待0.5s后接收数据
            time.sleep(1)
            serial_receiveData = [0 for x in range(40)]
            # 接收数组数据
            data = typeC_serial.read(40)
            serial_receiveData = [hex(i) for i in data]
            for i in range(len(serial_receiveData)):
                serial_receiveData[i] = int(serial_receiveData[i],16)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                self.result_signal.emit('未接收到正确数据，再次接收！\n')
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                self.isPassMFK &= False
                isThisPass = False
                break
            elif dataLen == 8:#成功
                if trueData[7] == int(state):#测试通过
                    self.isPassMFK &= True
                    isThisPass = True
                    break
                else:#测试不通过
                    self.isPassMFK &= False
                    isThisPass = False
                    break
            elif dataLen == 7:  # 失败
                self.isPassMFK &= False
                isThisPass = False
                break
            else:
                continue
        if not isThisPass:
            self.messageBox_signal.emit(['测试警告', 'MFK测试未通过，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()

        # 关闭串口
        typeC_serial.close()

    # 掉电保存测试
    def CPU_powerOffSaveTest(self, serial_transmitData: list):
        isThisPass = True
        try:
            # 打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告', f'Failed to open serial port。\n'])
            self.isPassPowerOffSave = False
            isThisPass = False
            self.cancelAllTest()
        loopStartTime = time.time()
        # for serial_transmitData in serial_transmitData_array:
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                isThisPass = False
                self.isPassPowerOffSave &= False
                break
            # 发送数据
            typeC_serial.write(bytes(serial_transmitData))
            # 等待0.5s后接收数据
            time.sleep(1)
            serial_receiveData = [0 for x in range(40)]
            # 接收数组数据
            data = typeC_serial.read(40)
            serial_receiveData = [hex(i) for i in data]
            for i in range(len(serial_receiveData)):
                serial_receiveData[i] = int(serial_receiveData[i], 16)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                isThisPass = False
                self.isPassPowerOffSave &= False
                break
            elif dataLen == 7:
                if trueData[6] == 0x00:  # 成功
                    isThisPass = True
                    self.isPassPowerOffSave &= True
                    break
                elif trueData[6] == 0x01:  # 失败
                    isThisPass = False
                    self.isPassPowerOffSave &= False
                    break
            else:
                continue
            # if not self.isPassPowerOffSave:
            #     break
        if not isThisPass:
            self.messageBox_signal.emit(['测试警告', '掉电保存测试不通过，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()


        # 关闭串口
        typeC_serial.close()

    # RTC测试
    def CPU_RTCTest(self,wr:str):
        isThisPass = True
        serial_transmitData=[]
        if wr == 'write':
            now = datetime.datetime.now()
            serial_transmitData = [0xAC, 13, 0x00, 0x07, 0x10, 0x00,
                                        now.year-2000, now.month, now.day,
                                        now.hour, now.minute, now.second,
                                        now.weekday(),self.getCheckNum( [0xAC, 13, 0x00, 0x07, 0x10, 0x00,
                                        now.year-2000, now.month, now.day,
                                        now.hour, now.minute, now.second,
                                        now.weekday()])]
        elif wr == 'read':
            serial_transmitData = [0xAC, 6, 0x00, 0x07, 0x0E, 0x00,
                                   self.getCheckNum([0xAC, 6, 0x00, 0x07, 0x0E, 0x00])]
        elif wr == 'battery':
            serial_transmitData = [0xAC, 6, 0x00, 0x07, 0x0E, 0x01,
                                   self.getCheckNum([0xAC, 6, 0x00, 0x07, 0x0E, 0x01])]
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
            self.isPassRTC = False
            isThisPass = False
            self.cancelAllTest()
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassRTC &= False
                isThisPass = False
                break
            now = datetime.datetime.now()
            # 发送数据
            typeC_serial.write(bytes(serial_transmitData))
            # 等待0.5s后接收数据
            time.sleep(1)
            serial_receiveData = [0 for x in range(40)]
            # 接收数组数据
            data = typeC_serial.read(40)
            serial_receiveData = [hex(i) for i in data]
            for i in range(len(serial_receiveData)):
                serial_receiveData[i] = int(serial_receiveData[i],16)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                self.result_signal.emit('未接收到正确数据，再次接收！\n')
                continue
            if dataLen == 3:  # 指令出错
                self.isPassRTC &= False
                isThisPass = False
                self.orderError(trueData[2])
                break
            elif trueData[4] == 0x10: #写
                if trueData[6] == 0x00:  # 写成功
                    self.isPassRTC &= True
                    isThisPass = True
                    week_dict={0:'一',1:'二',2:'三',3:'四',4:'五',5:'六',6:'日',}
                    self.result_signal.emit(f'写入时间为：{serial_transmitData[6]+2000}年{serial_transmitData[7]}月'
                                            f'{serial_transmitData[8]}日{serial_transmitData[9]}时'
                                            f'{serial_transmitData[10]}分{serial_transmitData[11]}秒，'
                                            f'星期{week_dict[serial_transmitData[12]]}\n')
                    break
                elif trueData[6] == 0x01:  # 写失败
                    self.isPassRTC &= False
                    isThisPass = False
                    break
            elif dataLen == 14:#读
                if trueData[6] == 0x00:  # 读成功
                    if (now.year-2000) == trueData[7] and now.month == trueData[8] and\
                        now.day == trueData[9] and now.hour == trueData[10] and\
                        now.weekday() == trueData[13]:
                        if (now.hour-trueData[10]) *3600+(now.minute-trueData[11])*60\
                                +(now.second-trueData[12])<=2:
                            self.isPassRTC &= True
                            isThisPass = True
                            week_dict = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日', }
                            # now1 = datetime.datetime.now()
                            self.result_signal.emit(f'当前北京时间为：{now.year}年{now.month}月{now.day}日'
                                             f'{now.hour}时{now.minute}分{now.second}秒，'
                                             f'星期{week_dict[now.weekday()]}\n')
                            self.result_signal.emit(f'读取PLC时间为：{trueData[7]+2000}年{trueData[8]}月{trueData[9]}日'
                                             f'{trueData[10]}时{trueData[11]}分{trueData[12]}秒，'
                                             f'星期{week_dict[trueData[13]]},合格\n')
                            break
                        else:
                            self.isPassRTC &= False
                            isThisPass = False
                            week_dict = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日', }
                            # now1 = datetime.datetime.now()
                            self.result_signal.emit(f'当前北京时间为：{now.year}年{now.month}月{now.day}日'
                                             f'{now.hour}时{now.minute}分{now.second}秒，'
                                             f'星期{week_dict[now.weekday()]}\n')

                            self.result_signal.emit(f'读取PLC时间为：{trueData[7] + 2000}年{trueData[8]}月{trueData[9]}日'
                                             f'{trueData[10]}时{trueData[11]}分{trueData[12]}秒，'
                                             f'星期{week_dict[trueData[13]]},不合格\n')
                            break
                    else:
                        self.isPassRTC &= False
                        isThisPass = False
                        self.result_signal.emit("时间读取失败\n")
                        break
                elif trueData[6] == 0x01:  # 读失败
                    self.isPassRTC &= False
                    isThisPass = False
                    self.result_signal.emit("时间读取失败\n")
                    break
            elif trueData[5] == 0x01:#纽扣电池状态
                if trueData[6] == 0x00:
                    self.isPassRTC &= True
                    isThisPass = True
                    if trueData[7] == 0x00:
                        self.result_signal.emit('纽扣电池未安装。\n')
                    elif trueData[7] == 0x01:
                        self.result_signal.emit('纽扣电池已安装。\n')
                    break
                elif trueData[6] == 0x01:
                    self.result_signal.emit('纽扣电池状态读取失败')
                    self.isPassRTC &= False
                    isThisPass = False
                    break
            else:
                continue
        if not isThisPass:
            self.messageBox_signal.emit(['测试警告', 'RTC测试不通过，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    # 各指示灯测试
    def CPU_LEDTest(self, transmitData:list,LED_name:str):
        isThisPass = True
        try:
            try:
                #打开串口
                typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
            except serial.SerialException as e:
                self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
                self.isPassLED = False
                isThisPass = False
                self.cancelAllTest()
            loopStartTime = time.time()
            while True:
                if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                    self.isPassLED &= False
                    isThisPass = False
                    break
                # 发送数据
                typeC_serial.write(bytes(transmitData))
                # 等待0.5s后接收数据
                time.sleep(0.1)
                serial_receiveData = [0 for x in range(40)]
                # 接收数组数据
                data = typeC_serial.read(40)
                serial_receiveData = [hex(i) for i in data]
                for i in range(len(serial_receiveData)):
                    serial_receiveData[i] = int(serial_receiveData[i],16)
                trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
                if isSendAgain:
                    self.result_signal.emit('未接收到正确数据，再次接收！\n')
                    continue
                if dataLen == 3:  # 指令出错
                    self.orderError(trueData[2])
                    self.isPassLED &= False
                    isThisPass = False
                    break
                if (dataLen == 7 and trueData[5] == 0x53) or (dataLen == 7 and trueData[5] == 0x50):
                    if trueData[6] == 0x00:  # 进入（退出）测试模式成功
                        break
                    elif trueData[6] == 0x01:  # 进入（退出）测试模式失败
                        self.isPassLED &= False
                        isThisPass = False
                        break
                elif dataLen == 7 and trueData[5] != 0xAA: #单个灯写
                    if trueData[6] == 0x00:  # 写入成功
                        self.isPassLED &= True
                        isThisPass = True
                    elif trueData[6] == 0x01:  # 写入失败
                        self.isPassLED &= False
                        isThisPass = False
                    break
                elif dataLen == 7 and trueData[5] == 0xAA: #所有灯写/读
                    if trueData[6] == 0x00:  # 写/读所有灯成功
                        self.isPassLED &= True
                        isThisPass = True
                    elif trueData[6] == 0x01:  # 写/读所有灯失败
                        self.isPassLED &= False
                        isThisPass = False
                    break
                elif dataLen == 8: #读
                    if trueData[7] == 0x00:#灯熄灭
                        # self.messageBox_signal.emit(['操作提示',f'{LED_name}是否熄灭？'])
                        # reply = self.result_queue.get()
                        # if reply == QMessageBox.AcceptRole:
                        #     self.isPassLED &= True
                        #     isThisPass = True
                        # else:
                        #     self.isPassLED &= False
                        #     isThisPass = False
                        break
                    elif trueData[7] == 0x01:#灯点亮
                        # image = self.CPU_pic_dir+f'/{LED_name}.png'
                        # self.pic_messageBox_signal.emit(['操作提示',f'{LED_name}是否如图所示点亮？',image])
                        # reply = self.result_queue.get()
                        # if reply == QMessageBox.AcceptRole:
                        #     self.isPassLED &= True
                        #     isThisPass = True
                        # else:
                        #     self.isPassLED &= False
                        #     isThisPass = False
                        break
                elif dataLen == 9:#读所有灯状态
                    if trueData[7] == 0x00 and trueData[8] == 0x00:
                        image = self.CPU_pic_dir + f'/{LED_name}.png'
                        self.pic_messageBox_signal.emit(['操作提示', f'{LED_name}（除PWR灯以外）是否如图所示熄灭？',image])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.AcceptRole:
                            self.isPassLED &= True
                            isThisPass = True
                        else:
                            self.isPassLED &= False
                            isThisPass = False
                        break
                    else:
                        self.isPassLED &= False
                        isThisPass = False
                        break
                else:
                    continue
        except Exception as e:
            self.result_signal.emit(f'LED error:{traceback.format_exc()}\n')
            self.isPassLED &= False
            isThisPass = False
        finally:
            if not isThisPass:
                self.messageBox_signal.emit(['测试警告', 'LED测试不通过，是否取消后续测试？'])
                reply = self.result_queue.get()
                if reply == QMessageBox.AcceptRole:
                    self.cancelAllTest()

            # 关闭串口
            typeC_serial.close()

    def CPU_LED_light_loop(self):
        w_transmitData = [0xAC, 7, 0x00, 0x09, 0x10, 0x00, 0x01,
                          self.getCheckNum([0xAC, 7, 0x00, 0x09, 0x10, 0x00, 0x01])]
        r_transmitData = [0xAC, 6, 0x00, 0x09, 0x0E, 0x00,
                          self.getCheckNum([0xAC, 6, 0x00, 0x09, 0x0E, 0x00])]
        name_LED = ['RUN灯', 'ERROR灯', 'BAT_LOW灯', 'PRG灯', 'RS-232C灯', 'RS-485灯', 'HOST灯']
        for l in range(7):
            if self.stopChannel == QMessageBox.AcceptRole or self.stopChannel == QMessageBox.RejectRole:
                self.stopChannel= 2
                # 打开所有LED
                w_allLED = [0xAC, 8, 0x00, 0x09, 0x10, 0xAA, 0x01, 0x00, 0x88]
                self.CPU_LEDTest(transmitData=w_allLED, LED_name='所有灯')
                break

            if not self.isCancelAllTest:
                w_transmitData[5] = l
                w_transmitData[7] = self.getCheckNum(w_transmitData[:7])
                self.CPU_LEDTest(transmitData=w_transmitData, LED_name=name_LED[l])
                time.sleep(0.2)
                if not self.isCancelAllTest:
                    # 关闭所有LED
                    w_allLED = [0xAC, 8, 0x00, 0x09, 0x10, 0xAA, 0x00, 0x00, 0x88]
                    self.CPU_LEDTest(transmitData=w_allLED, LED_name='所有灯')
                    if l == 6:
                        l = 0
        pass

    def CPU_InTest(self, transmitData:list):
        isThisPass = True
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
            self.isPassIn = False
            isThisPass = False
            self.cancelAllTest()
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassIn &= False
                isThisPass = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            # 等待0.5s后接收数据
            time.sleep(1)
            serial_receiveData = [0 for x in range(40)]
            # 接收数组数据
            data = typeC_serial.read(40)
            serial_receiveData = [hex(i) for i in data]
            for i in range(len(serial_receiveData)):
                serial_receiveData[i] = int(serial_receiveData[i],16)
            trueData,dataLen,isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                self.result_signal.emit('未接收到正确数据，再次接收！\n')
                continue
            if dataLen == 3:#指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 12:#读所有输入通道
                if trueData[7] == 0x00:  # 读所有输入通道成功
                    if trueData[9] == 0xAA and trueData[10] == 0xAA and trueData[11] == 0xAA:
                        image_aa = self.CPU_pic_dir + '/AAAA.png'
                        self.pic_messageBox_signal.emit(['操作提示', 'CPU输入通道指示灯是否如图所示点亮？', image_aa])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.AcceptRole:
                            self.isPassIn &= True
                            isThisPass = True
                        else:
                            self.result_signal.emit('CPU输入通道或模块QN0016存在故障。\n')
                            self.isPassIn &= False
                            isThisPass = False
                        break
                    elif trueData[9] == 0x55 and trueData[10] == 0x55 and trueData[11] == 0x55:
                        image_55 = self.CPU_pic_dir+'/5555.png'
                        self.pic_messageBox_signal.emit(['操作提示','CPU输入通道指示灯是否如图所示点亮？',image_55])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.AcceptRole:
                            self.isPassIn &= True
                            isThisPass = True
                        else:
                            self.result_signal.emit('CPU输入通道或模块QN0016存在故障。\n')
                            self.isPassIn &= False
                            isThisPass = False
                        break
                    elif trueData[9] == 0x00 and trueData[10] == 0x00 and trueData[11] == 0x00:
                        image_00 = self.CPU_pic_dir + '/0000.png'
                        self.pic_messageBox_signal.emit(['操作提示', 'CPU输入通道指示灯是否如图所示全部熄灭？', image_00])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.AcceptRole:
                            self.isPassIn &= True
                            isThisPass = True
                        else:
                            self.result_signal.emit('CPU输入通道或模块QN0016存在故障。\n')
                            self.isPassIn &= False
                            isThisPass = False
                        break
                    else:
                        # ch1_8 = trueData[9][:2]+trueData[9][]
                        hex_trueData = [hex(i) for i in trueData]
                        self.in_bin_str = [(bin(int(hex_trueData[9], 16))[2:])[::-1],#颠倒字符串顺序
                                           (bin(int(hex_trueData[10], 16))[2:])[::-1],
                                            (bin(int(hex_trueData[11], 16))[2:])[::-1]]
                        for ibs in range(3):
                            if len(self.in_bin_str[ibs])<8:
                                for ii in range(8-len(self.in_bin_str[ibs])):
                                    self.in_bin_str[ibs] +='0'
                        self.inCh_bin = self.in_bin_str[0]+self.in_bin_str[1]+self.in_bin_str[2]
                        self.result_signal.emit(f'CPU输入通道1-8:{self.in_bin_str[0]}\n'
                                                f'CPU输入通道9-16:{self.in_bin_str[1]}\n'
                                                f'CPU输入通道17-24:{self.in_bin_str[2]}\n')
                        self.isPassIn &= False
                        isThisPass = False
                        break
                elif trueData[7] == 0x01:  # 读所有输入通道失败
                    self.isPassIn &= False
                    isThisPass = False
                    break

        if not isThisPass:
            self.messageBox_signal.emit(['测试警告', '本体输入测试不通过，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()

        # 关闭串口
        typeC_serial.close()

    #本体OUT测试
    def CPU_OutTest(self, transmitData:list):
        isThisPass = True
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
            self.isPassOut = False
            isThisPass = False
            self.cancelAllTest()
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassOut &= False
                isThisPass = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            # 等待0.5s后接收数据
            time.sleep(1)
            serial_receiveData = [0 for x in range(40)]
            # 接收数组数据
            data = typeC_serial.read(40)
            serial_receiveData = [hex(i) for i in data]
            for i in range(len(serial_receiveData)):
                serial_receiveData[i] = int(serial_receiveData[i],16)
            trueData,dataLen,isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                self.result_signal.emit('未接收到正确数据，再次接收！\n')
                continue
            if dataLen == 3:#指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 8:#读所有输出通道
                if trueData[7] == 0x00:  # 读所有输出通道成功
                    time.sleep(0.2)
                    tr = time.time()
                    self.clearList(self.m_receiveData)
                    while True:
                        bool_receive, self.m_can_obj = CAN_option.receiveCANbyID((0x180 + self.CANAddr1), 2000,1)

                        self.m_receiveData = self.m_can_obj.Data
                        # self.result_signal.emit(
                        #     f'{hex(self.m_receiveData[0])},{hex(self.m_receiveData[1])},{hex(self.m_receiveData[2])}'
                        #     f',{hex(self.m_receiveData[3])},{hex(self.m_receiveData[4])},{hex(self.m_receiveData[5])}'
                        #     f',{hex(self.m_receiveData[6])},{hex(self.m_receiveData[7])}\n\n')
                        if (time.time() - tr)*1000 > 300:
                            break
                    if bool_receive == False:
                        self.messageBox_signal.emit(['警告','ET1600接收数据超时，请检查ET1600接线！'])
                        self.isPassOut &= False
                        isThisPass = False
                        break
                    elif (hex(self.m_receiveData[0]) == '0xaa' and hex(self.m_receiveData[1]) == '0xaa') or \
                            (hex(self.m_receiveData[0]) == '0x55' and hex(self.m_receiveData[1]) == '0x55') or \
                            (hex(self.m_receiveData[0]) == '0x0' and hex(self.m_receiveData[1]) == '0x0'):
                        if hex(self.m_receiveData[0]) == '0xaa':
                            image_aa = self.CPU_pic_dir + '/AE_aa.png'
                            self.pic_messageBox_signal.emit(['操作提示','模块ET1600的通道指示灯是否如图所示点亮？',image_aa])
                        if hex(self.m_receiveData[0]) == '0x55':
                            image_55 = self.CPU_pic_dir + '/AE_55.png'
                            self.pic_messageBox_signal.emit(['操作提示','模块ET1600的通道指示灯是否如图所示点亮？',image_55])
                        elif hex(self.m_receiveData[0]) == '0x0':
                            image_00 = self.CPU_pic_dir + '/AE_00.png'
                            self.pic_messageBox_signal.emit(
                                ['操作提示', '模块ET1600的通道指示灯是否如图所示全部熄灭？', image_00])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.AcceptRole:
                            self.isPassOut &= True
                            isThisPass = True
                        else:
                            self.result_signal.emit('CPU输出通道或模块ET1600存在故障。\n')
                            self.isPassOut &= False
                            isThisPass = False
                        break
                    else:#读到的ET1600通道数据不对
                        self.isPassOut &= False
                        isThisPass = False
                        if hex(transmitData[8]) == '0xff':
                            self.result_signal.emit(f'{hex(self.m_receiveData[0])}\n{hex(self.m_receiveData[1])}\n')
                            self.out_bin_str = [(bin(int(hex(self.m_receiveData[0]), 16))[2:])[::-1],
                                                (bin(int(hex(self.m_receiveData[1]), 16))[2:])[::-1]]
                            self.outCh_bin = self.out_bin_str[0] + self.out_bin_str[1]
                            self.result_signal.emit(f'ET1600通道1-8 = {self.outCh_bin[:8]}\n'
                                                    f'ET1600通道9-16 = {self.outCh_bin[8:]}\n')
                            for obs in range(2):
                                if len(self.out_bin_str[obs]) < 8:
                                    for ii in range(8 - len(self.out_bin_str[obs])):
                                        self.out_bin_str[obs] += '0'
                        # elif transmitData[8] == '0x0':
                        #     self.out_bin_str = [''.join(['1' if b == '0' else '0' for b in (bin(int(hex(self.m_receiveData[0]), 16))[2:])]),
                        #                         ''.join(['1' if b == '0' else '0' for b in (bin(int(hex(self.m_receiveData[1]), 16))[2:])]),]
                elif trueData[7] == 0x01:  # 写所有输出通道失败
                    self.result_signal.emit('CPU输出通道写入失败。\n')
                    self.isPassOut &= False
                    isThisPass = False
                    break
            else:
                continue

        if not isThisPass:
            self.messageBox_signal.emit(['测试警告', '本体输出测试不通过，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    def TCP_UDPTest(self):
        isThisPass = True
        try:
            import socket
            # 创建TCP连接
            client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            server_address = ('192.168.1.66', 8000)
            client_socket.connect(server_address)

            for i in range(100):
                # 发送数据
                message = f'------TCP收发测试------TCP收发测试------TCP收发测试------'
                client_socket.send(message.encode())
                if i== 99:
                    self.result_signal.emit(f'TCP发送的数据：{message}')

                # 接收数据
                data = client_socket.recv(1024)
                if i == 99:
                    self.result_signal.emit(f'TCP接收的数据：{data.decode()}')

            # 关闭TCP连接
            client_socket.close()

            # 创建UDP连接
            client_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            client_address = ('192.168.1.100', 1025)
            client_socket.bind(client_address)
            for i in range(100):
                # 发送数据
                server_address = ('192.168.1.66', 1025)
                message = f'------UDP收发测试------UDP收发测试------UDP收发测试------'
                client_socket.sendto(message.encode(), server_address)
                if i == 99:
                    self.result_signal.emit(f'UDP发送的数据：{message}')

                # 接收数据
                data, address = client_socket.recvfrom(1024)
                if i == 99:
                    self.result_signal.emit(f'UDP接收的数据：{data.decode()}')

            # 关闭UDP连接
            client_socket.close()
            self.isPassETH &= True
            isThisPass = True
        except:
            self.isPassETH &= False
            isThisPass = False
        finally:
            if not isThisPass:
                self.messageBox_signal.emit(['测试警告', f'本体ETH测试不通过，是否取消后续测试？'])
                reply = self.result_queue.get()
                if reply == QMessageBox.AcceptRole:
                    self.cancelAllTest()


    # 本体ETH测试
    def CPU_ETHTest(self, transmitData: list,isPassSign:bool):
        isThisPass = True
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassETH &= False
                isThisPass = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            # 等待0.5s后接收数据
            time.sleep(1)
            serial_receiveData = [0 for x in range(40)]
            # 接收数组数据
            data = typeC_serial.read(40)
            serial_receiveData = [hex(i) for i in data]
            for i in range(len(serial_receiveData)):
                serial_receiveData[i] = int(serial_receiveData[i],16)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                self.result_signal.emit('未接收到正确数据，再次接收！\n')
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 11:  # 读IP
                if trueData[6] == 0x00:  # 读IP成功
                    if trueData[7] == transmitData[5] and trueData[8] == transmitData[6] \
                            and trueData[9] == transmitData[7] and trueData[10] == transmitData[8]:
                        isPassSign &= True
                        isThisPass = True
                        break
                    else:
                        isPassSign &= False
                        isThisPass = False
                        break
                elif trueData[6] == 0x01:  # 读IP失败
                    isPassSign &= False
                    isThisPass = False
                    break
            elif dataLen == 9:  # 读UDP
                if trueData[6] == 0x00:  # 读UDP成功
                    if trueData[7] == transmitData[5] and trueData[8] == transmitData[6]:
                        isPassSign &= True
                        isThisPass = True
                        break
                    else:
                        isPassSign &= False
                        isThisPass = False
                        break
                elif trueData[6] == 0x01:  # 读UDP失败
                    isPassSign &= False
                    isThisPass = False
                    break
            else:
                continue

        if not self.isPassETH_IP and self.isPassETH_UDP:
            self.messageBox_signal.emit(['测试警告', '本体ETH IP测试不通过，是否取消后续测试？'])
            self.isPassETH = False
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()
        elif (not self.isPassETH_UDP and self.isPassETH_IP) or (not self.isPassETH_UDP and not self.isPassETH_IP):
            self.isPassETH = False
            self.messageBox_signal.emit(['测试警告', '本体ETH UDP测试不通过，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    # 本体232测试
    def CPU_232Test(self, transmitData: list,baudRate:list):
        isThisPass = True
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
            self.isPass232 = False
            isThisPass = False
            self.cancelAllTest()
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPass232 &= False
                isThisPass = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            # 等待0.5s后接收数据
            time.sleep(1)
            serial_receiveData = [0 for x in range(40)]
            # 接收数组数据
            data = typeC_serial.read(40)
            serial_receiveData = [hex(i) for i in data]
            for i in range(len(serial_receiveData)):
                serial_receiveData[i] = int(serial_receiveData[i],16)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                self.result_signal.emit('未接收到正确数据，再次接收！\n')
                self.isPass232 &= False
                isThisPass = False
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                break
            elif dataLen == 7:  # 写
                if trueData[6] == 0x00:  # 写成功
                    self.isPass232 &= True
                    isThisPass = True
                    break
                elif trueData[6] == 0x01:  # 写失败
                    self.isPass232 &= False
                    isThisPass = False
                    break
            elif dataLen == 11:  # 读
                if trueData[6] == 0x00:  # 读成功
                    if trueData[7] == baudRate[0] and trueData[8] == baudRate[1]\
                        and trueData[9] == baudRate[2] and trueData[10] == baudRate[3]:
                        self.isPass232 &= True
                        isThisPass = True
                    else:
                        self.isPass232 &= False
                        isThisPass = False
                    break
                elif trueData[6] == 0x01:  # 读失败
                    self.isPass232 &= False
                    isThisPass = False
                    break
            else:
                continue
        if not isThisPass:
            self.messageBox_signal.emit(['测试警告', '本体232测试不通过，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    def transmitBy232or485(self,type:str,baudRate:int,transmitData:list):
        isReceiveTrueData:bool = True
        try:
            m_port = ' '
            #打开串口
            if type =='232':
                m_port = str(self.serialPort_232)
                typeC_serial = serial.Serial(port=m_port, baudrate=baudRate, timeout=1)
            elif type =='485':
                m_port = str(self.serialPort_485)
                typeC_serial = serial.Serial(port=m_port, baudrate=baudRate, timeout=1)
            elif type == 'power':
                m_port = str(self.serialPort_power)
                typeC_serial = serial.Serial(port=m_port, baudrate=baudRate, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{m_port}打开失败，请检查该串口是否被占用。\n'])
            isReceiveTrueData &= False
            # 关闭串口
            typeC_serial.close()
            return isReceiveTrueData

        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                isReceiveTrueData = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            hex_transmitData = [hex(ht) for ht in transmitData]
            self.result_signal.emit(f'发送的数据：{hex_transmitData}')
            # 等待0.5s后接收数据
            time.sleep(1)
            serial_receiveData = [0 for x in range(len(transmitData))]
            # 接收数组数据
            data = typeC_serial.read(len(transmitData))
            if len(data) == 0:
                self.messageBox_signal.emit(['操作警告',f'未接收到{type}信号，请检查{type}接线是否断开。'])
                reply = self.result_queue.get()
                if reply == QMessageBox.AcceptRole:
                    isReceiveTrueData &= False
                else:
                    isReceiveTrueData &= False
                break
            serial_receiveData = [hex(i) for i in data]
            # for i in range(len(serial_receiveData)):
            #     serial_receiveData[i] = int(serial_receiveData[i], 16)
            for tt in range(len(transmitData)):
                if hex(transmitData[tt]) != serial_receiveData[tt]:
                    isReceiveTrueData &=False
                    self.result_signal.emit(f'接收的数据：{serial_receiveData}，错误。\n\n')
                    break
            if isReceiveTrueData:
                self.result_signal.emit(f'接收的数据：{serial_receiveData}，正确。\n\n')
                isReceiveTrueData &=True
                break


        if not isReceiveTrueData:
            self.messageBox_signal.emit(['测试警告', '数据接收失败，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()
        return isReceiveTrueData

    def powerControl(self,baudRate:int,transmitData:list):
        isReceiveTrueData:bool = True
        try:
            m_port = ' '
            #打开串口
            m_port = str(self.serialPort_power)
            typeC_serial = serial.Serial(port=m_port, baudrate=baudRate, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告', f'串口{m_port}打开失败，请检查该串口是否被占用。\n'])
            isReceiveTrueData &= False
            # 关闭串口
            typeC_serial.close()
            return isReceiveTrueData

        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                isReceiveTrueData = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            hex_transmitData = [hex(ht) for ht in transmitData]
            self.result_signal.emit(f'发送的数据：{hex_transmitData}')
            # 等待0.5s后接收数据
            time.sleep(1)
            serial_receiveData = [0 for x in range(len(transmitData))]
            # 接收数组数据
            ## data = typeC_serial.read(len(transmitData))
            # if len(data) == 0:
            #     self.messageBox_signal.emit(['操作警告', '未接收到信号，请检查可编程电源的通信线是否断开。'])
            #     reply = self.result_queue.get()
            #     if reply == QMessageBox.AcceptRole:
            #         isReceiveTrueData &= False
            #     else:
            #         isReceiveTrueData &= False
            #     break
            # serial_receiveData = [hex(i) for i in data]
            # # for i in range(len(serial_receiveData)):
            # #     serial_receiveData[i] = int(serial_receiveData[i], 16)
            # for tt in range(len(transmitData)):
            #     if hex(transmitData[tt]) != serial_receiveData[tt]:
            #         isReceiveTrueData &= False
            #         self.result_signal.emit(f'接收的数据：{serial_receiveData}，错误。\n\n')
            #         break
            # if isReceiveTrueData:
            #     self.result_signal.emit(f'接收的数据：{serial_receiveData}，正确。\n\n')
            #     isReceiveTrueData &= True
            #     break

        # if not isReceiveTrueData:
        #     self.messageBox_signal.emit(['测试警告', '控制可编程电源失败，是否取消后续测试？'])
        #     reply = self.result_queue.get()
        #     if reply == QMessageBox.AcceptRole:
        #         self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()


    # 本体485测试
    def CPU_485Test(self, transmitData: list,baudRate:list):
        isThisPass = True
        try:
            # 打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(
                ['错误警告', f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
            self.isPass485 = False
            isThisPass = False
            self.cancelAllTest()
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPass485 &= False
                isThisPass = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            # 等待0.5s后接收数据
            time.sleep(1)
            serial_receiveData = [0 for x in range(40)]
            # 接收数组数据
            data = typeC_serial.read(40)
            serial_receiveData = [hex(i) for i in data]
            for i in range(len(serial_receiveData)):
                serial_receiveData[i] = int(serial_receiveData[i], 16)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                self.result_signal.emit('未接收到正确数据，再次接收！\n')
                self.isPass485 &= False
                isThisPass = False
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                self.isPass485 = False
                isThisPass = False
                break
            elif dataLen == 7:  # 写
                if trueData[6] == 0x00:  # 写成功
                    self.isPass485 &= True
                    isThisPass = True
                    break
                elif trueData[6] == 0x01:  # 写失败
                    self.isPass485 &= False
                    isThisPass = False
                    break
            elif dataLen == 11:  # 读
                if trueData[6] == 0x00:  # 读成功
                    if trueData[7] == baudRate[0] and trueData[8] == baudRate[1] \
                            and trueData[9] == baudRate[2] and trueData[10] == baudRate[3]:
                        self.isPass485 &= True
                        isThisPass = True
                    else:
                        self.isPass485 &= False
                        isThisPass = False
                    break
                elif trueData[6] == 0x01:  # 读失败
                    self.isPass485 &= False
                    isThisPass = False
                    break
            else:
                continue
        if not isThisPass:
            self.messageBox_signal.emit(['测试警告', '本体485测试不通过，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    # 本体右扩CAN测试
    def CPU_rightCANTest(self, transmitData: list,justOff:bool=False):
        isThisPass = True
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
            self.isPassRightCAN =False
            isThisPass = False
            self.cancelAllTest()
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassRightCAN &= False
                isThisPass = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            if justOff:
                break
            # 等待0.5s后接收数据
            time.sleep(0.5)
            serial_receiveData = [0 for x in range(40)]
            # 接收数组数据
            data = typeC_serial.read(40)
            serial_receiveData = [hex(i) for i in data]
            for i in range(len(serial_receiveData)):
                serial_receiveData[i] = int(serial_receiveData[i],16)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                self.result_signal.emit('未接收到正确数据，再次接收！\n')
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                self.isPassRightCAN &= False
                isThisPass = False
                break
            elif dataLen == 7:  # 写
                if trueData[6] == 0x00:  # 写成功
                    if transmitData[5] == 0x00 and transmitData[6] == 0x00:#复位灯熄灭
                        self.messageBox_signal.emit(['操作提示','右扩CAN的RESET灯（绿灯）是否熄灭?'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.AcceptRole:
                            self.isPassRightCAN &= True
                            isThisPass = True
                        else:
                            self.isPassRightCAN &= False
                            isThisPass = False
                    elif transmitData[5] == 0x00 and transmitData[6] == 0x01:#复位灯亮
                        self.messageBox_signal.emit(['操作提示','右扩CAN的RESET灯（绿灯）是否点亮?'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.AcceptRole:
                            self.isPassRightCAN &= True
                            isThisPass = True
                        else:
                            self.isPassRightCAN &=False
                            isThisPass = False
                    elif transmitData[5] == 0x01 and transmitData[6] == 0x01:#使能灯灭
                        self.messageBox_signal.emit(['操作提示','右扩CAN的使能灯（红灯）是否熄灭?'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.AcceptRole:
                            self.isPassRightCAN &= True
                            isThisPass = True
                        else:
                            self.isPassRightCAN &=False
                            isThisPass = False
                    elif transmitData[5] == 0x01 and transmitData[6] == 0x00:#使能灯亮
                        self.messageBox_signal.emit(['操作提示','右扩CAN的使能灯（红灯）是否点亮?'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.AcceptRole:
                            self.isPassRightCAN &= True
                            isThisPass = True
                        else:
                            self.isPassRightCAN &=False
                            isThisPass = False
                    break
                elif trueData[6] == 0x01:  # 写失败
                    self.isPassRightCAN &= False
                    isThisPass = False
                    break
            else:
                continue
        if not isThisPass:
            self.messageBox_signal.emit(['测试警告', '本体右扩CAN测试不通过，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    #MAC地址和3码
    def CPU_configTest(self, transmitData: list):
        isThisPass = True
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
            self.isPassConfig = False
            isThisPass = False
            self.cancelAllTest()
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassConfig &= False
                isThisPass = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            # 等待0.5s后接收数据
            time.sleep(1)
            serial_receiveData = [0 for x in range(40)]
            # 接收数组数据
            data = typeC_serial.read(40)
            serial_receiveData = [hex(i) for i in data]
            for i in range(len(serial_receiveData)):
                serial_receiveData[i] = int(serial_receiveData[i],16)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            # self.result_signal.emit(f'trueData:{trueData}')
            if isSendAgain:
                self.result_signal.emit('未接收到正确数据，再次接收！\n')
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                self.isPassConfig &= False
                isThisPass = False
                break
            elif dataLen == 7:  # 写
                if trueData[6] == 0x00:  # 写成功
                    self.isPassConfig &= True
                    isThisPass = True
                    break
                elif trueData[6] == 0x01:  # 写/读失败
                    self.isPassConfig &= False
                    isThisPass = False
                    break
            elif trueData[5] == 0x00:  # 读MAC
                if trueData[6] == 0x00:  # 读成功
                    if trueData[7] == 6:
                        MAC_int = ''
                        for MACNum in range(6):
                            mac_stand = self.MAC_array
                            if (hex(trueData[MACNum+8]) != hex(mac_stand[MACNum])):
                                self.isPassConfig &= False
                                isThisPass = False
                                self.result_signal.emit('MAC读取失败\n')
                                break
                            else:
                                MAC_int += f'{int(trueData[MACNum+8])}'
                                if MACNum<5:
                                    MAC_int +='-'
                                if MACNum == 5:
                                    self.result_signal.emit(f'MAC读取成功，MAC值为：{MAC_int}\n')
                                self.isPassConfig &= True
                                isThisPass = True
                    else:
                        self.isPassConfig &= False
                        isThisPass = False
                    break
                elif trueData[6] == 0x01:  # 读失败
                    self.isPassConfig = False
                    isThisPass = False
                    break
            elif trueData[5] == 0x02:  # 读SN
                if trueData[6] == 0x00:  # 读成功
                    if trueData[7] == 16:
                        SN_int = ''
                        for SNNum in range(12):
                            if (int(trueData[SNNum+8]) != ord(self.module_sn[SNNum])):
                                self.isPassConfig &= False
                                isThisPass = False
                                self.result_signal.emit('SN读取失败\n')
                            else:
                                SN_int += chr(trueData[SNNum+8])
                                self.isPassConfig &= True
                                isThisPass = True
                                if SNNum == 11:
                                    self.result_signal.emit(f'SN读取成功,SN值为：{SN_int}\n')
                    else:
                        self.isPassConfig &= False
                        isThisPass = False
                    break
                elif trueData[6] == 0x01:  # 读失败
                    self.isPassConfig &= False
                    isThisPass = False
                    break
            elif trueData[5] == 0x01:  # 读PN
                if trueData[6] == 0x00:  # 读成功
                    if trueData[7] == 24:
                        PN_int = ''
                        for PNNum in range(14):
                            if (int(trueData[PNNum+8]) != ord(self.module_pn[PNNum])):
                                self.isPassConfig &= False
                                isThisPass = False
                                self.result_signal.emit('PN读取失败\n')
                            else:
                                PN_int += chr(trueData[PNNum+8])
                                self.isPassConfig &= True
                                isThisPass = True
                                if PNNum == 13:
                                    self.result_signal.emit(f'PN读取成功，PN值为：{PN_int}\n')
                    else:
                        self.isPassConfig &= False
                        isThisPass = False
                    break
                elif trueData[6] == 0x01:  # 读失败
                    self.isPassConfig &= False
                    isThisPass = False
                    break
            elif trueData[5] == 0x03:  # 读REV
                if trueData[6] == 0x00:  # 读成功
                    if trueData[7] == 8:
                        REV_int = ''
                        for REVNum in range(2):
                            if (int(trueData[REVNum+8]) != ord(self.module_rev[REVNum])):
                                self.isPassConfig &= False
                                isThisPass = False
                                self.result_signal.emit('REV读取失败\n')
                            else:
                                REV_int += chr(trueData[REVNum+8])
                                self.isPassConfig &= True
                                isThisPass = True
                                if REVNum == 1:
                                    self.result_signal.emit(f'REV读取成功,REV值为：{REV_int}\n')
                    else:
                        self.isPassConfig &= False
                        isThisPass = False
                    break
                elif trueData[6] == 0x01:  # 读失败
                    self.isPassConfig &= False
                    isThisPass = False
                    break
            else:
                continue
        if not isThisPass:
            self.messageBox_signal.emit(['测试警告', 'MAC与三码写入测试不通过，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

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
                    ['错误警告', f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
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

    def l_distanceCAN(self):
        w_transmit = [0xac, 7,0x00,0x10,0x10,0x00,0x00,0x00]
        baudRate = [0x00,0x01,0x02,0x03,0x04,0x05,0x06,0x07]
        self.isPassLdCAN =True
        for ld in range(1,8):
            w_transmit[6] = ld
            w_transmit[7] = self.getCheckNum(w_transmit[:7])
            serial_transmitData = w_transmit
            isThisPass = True
            try:
                # 打开串口
                typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
            except serial.SerialException as e:
                self.messageBox_signal.emit(
                    ['错误警告', f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
                self.isPassLdCAN &= False
                isThisPass = False
                self.cancelAllTest()
            loopStartTime = time.time()
            while True:
                if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                    self.isPassLdCAN &= False
                    isThisPass = False
                    break
                now = datetime.datetime.now()
                # 发送数据
                typeC_serial.write(bytes(serial_transmitData))
                # 等待0.5s后接收数据
                time.sleep(1)
                serial_receiveData = [0 for x in range(40)]
                # 接收数组数据
                data = typeC_serial.read(40)
                serial_receiveData = [hex(i) for i in data]
                for i in range(len(serial_receiveData)):
                    serial_receiveData[i] = int(serial_receiveData[i], 16)
                trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
                if isSendAgain:
                    self.result_signal.emit('未接收到正确数据，再次接收！\n')
                    continue
                if dataLen == 3:  # 指令出错
                    self.orderError(trueData[2])
                    self.isPassLdCAN &= False
                    isThisPass = False
                    break
                elif trueData[4] == 0x10:
                    if trueData[6] == 0x00:
                        self.isPassLdCAN &= True
                        isThisPass = True
                    elif trueData[6] == 0x01:
                        self.isPassLdCAN &= False
                        isThisPass = False
                    break
                elif trueData[4] == 0x0E:
                    if trueData[6] == 0x00:
                        if trueData[7] == baudRate[ld]:
                            self.isPassLdCAN &= True
                            isThisPass = True
                        else:
                            self.isPassLdCAN &= False
                            isThisPass = False
                    elif trueData[6] == 0x01:
                        self.isPassLdCAN &= False
                        isThisPass = False
                    break
                else:
                    continue
            # 关闭串口
            typeC_serial.close()
        if not isThisPass:
            self.messageBox_signal.emit(['测试警告', '远程CAN测试不通过，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()
        else:
            self.result_signal.emit('远程CAN测试通过。\n')


    def DIP_SW(self):
        r_transmit = [0xac,6, 0x00, 0x11, 0x0e, 0x00,self.getCheckNum([0xac,6, 0x00, 0x11, 0x0e, 0x00])]
        serial_transmitData = r_transmit
        isThisPass = True
        self.isPassDIP = True
        try:
            # 打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(
                ['错误警告', f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'])
            self.isPassDIP &= False
            isThisPass = False
            self.cancelAllTest()
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassDIP &= False
                isThisPass = False
                break
            now = datetime.datetime.now()
            # 发送数据
            typeC_serial.write(bytes(serial_transmitData))
            # 等待0.5s后接收数据
            time.sleep(1)
            serial_receiveData = [0 for x in range(40)]
            # 接收数组数据
            data = typeC_serial.read(40)
            serial_receiveData = [hex(i) for i in data]
            for i in range(len(serial_receiveData)):
                serial_receiveData[i] = int(serial_receiveData[i], 16)
            trueData, dataLen, isSendAgain = self.dataJudgement(serial_receiveData)
            if isSendAgain:
                self.result_signal.emit('未接收到正确数据，再次接收！\n')
                continue
            if dataLen == 3:  # 指令出错
                self.orderError(trueData[2])
                self.isPassDIP &= False
                isThisPass = False
                break
            elif trueData[6] == 0x00:
                self.isPassDIP &= True
                isThisPass = True
                if trueData[7] == 0x00:
                    self.result_signal.emit('拨码未全OFF状态。\n')
                elif trueData[7] == 0x01:
                    self.result_signal.emit('拨码未全ON状态。\n')
                break
            elif trueData[6] == 0x01:
                self.isPassDIP &= False
                isThisPass = False
                break
            else:
                continue
        if not isThisPass:
            self.messageBox_signal.emit(['测试警告', 'DIP_SW测试不通过，是否取消后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.AcceptRole:
                self.cancelAllTest()
        else:
            self.result_signal.emit('DIP_SW测试通过。\n')
        # 关闭串口
        typeC_serial.close()

        #判断接收的数据是否正确
    def dataJudgement(self,serial_receiveData:list):
        isSendAgain = False
        trueData = [ ]
        dataLen = 0
        # 数据头位置
        startNum = 0
        hex_serial_receiveData = [hex(i) for i in serial_receiveData]
        # print(f'serial_receiveData：{hex_serial_receiveData}')
        # 确定数据头
        for i in range(len(serial_receiveData)):
            if serial_receiveData[i] == int(0xCA):
                # self.result_signal.emit(f'找到startNum:{startNum}\n')
                startNum = i
                isSendAgain = False
                break
            if i == len(serial_receiveData) - 1 and serial_receiveData[i] != int(0xCA):
                # self.result_signal.emit(f'未找到startNum\n')
                isSendAgain = True
                return trueData,dataLen,isSendAgain
        # self.result_signal.emit(f'serial_receiveData:{serial_receiveData}')
        # self.result_signal.emit(f'startNum:{startNum}')
        dataLen = int(serial_receiveData[startNum + 1])

        # 验证数据接收是否正确
        dataSum = 0
        for i in range(dataLen):
            dataSum += serial_receiveData[startNum + i]
        ckeckSum = self.getCheckNum(serial_receiveData[startNum:startNum+dataLen])
        # print(f"serial_receiveData[startNum:startNum+dataLen]:{serial_receiveData[startNum:startNum+dataLen]}")
        if ckeckSum == int(serial_receiveData[startNum + dataLen]):
            # self.result_signal.emit("接收的数据正确。\n")
            trueData = serial_receiveData[startNum:startNum + dataLen]
        hex_trueData = [hex(i) for i in trueData]
        if not self.is_running:
            return [], 0, True
        # self.result_signal.emit(f'receiveData:{hex_trueData}\n')
        return trueData,dataLen,isSendAgain

    #获取校验位
    def getCheckNum(self,data:list):
        sum = 0
        for d in data:
            sum += d
        checkNum = (~sum) & 0xff
        return checkNum

    def orderError(self,errCode):
        errorCode = {0x80: '消息长度错误', 0x81: '设备错误', 0x82: '测试类型错误',
                     0x83: '操作不支持', 0x84: '参数错误'}
        self.showErrorInf(f'指令发送错误：{hex(errCode)}，{errorCode[errCode]}\n'
                          f'取消剩余测试。\n\n')
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
        # ·
        self.label_signal.emit(['fail', '测试停止'])

    def showErrorInf(self,Inf):
        self.pauseOption()
        if not self.is_running:
            self.cancelAllTest()
            return
        self.result_signal.emit(f"{Inf}程序出现问题\n")
        # 捕获异常并输出详细的错误信息
        self.result_signal.emit(f"ErrorInf:\n{traceback.format_exc()}\n")
        # self.messageBox_signal.emit(['错误警告',(f"ErrorInf:\n{traceback.format_exc()}\n")])

    #CAN设备初始化
    def CAN_init(self,CAN_channel:list):
        CAN_option.close(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX)
        time.sleep(0.1)
        QApplication.processEvents()

        for channel in CAN_channel:
            # if channel == 0:
            #     #波特率 = 100k
            #     CAN_option.TIMING_0 = 0x04
            #     CAN_option.TIMING_1 = 0x1C
            # elif channel == 1:
            #     #波特率 = 1000k
            #     CAN_option.TIMING_0 = 0x00
            #     CAN_option.TIMING_1 = 0x14
            if not CAN_option.connect(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX, channel):
                # self.showMessageBox(['CAN设备存在问题', 'CAN设备开启失败，请检查CAN设备！'])
                # self.CANFail()
                return [False, ['CAN设备存在问题', f'CAN设备通道{channel}开启失败，请检查CAN设备！']]

            if not CAN_option.init(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX, channel, init_config):
                # self.showMessageBox(['CAN设备存在问题', 'CAN通道初始化失败，请检查CAN设备！'])
                # self.CANFail()
                return [False, ['CAN设备存在问题', f'CAN通道{channel}初始化失败，请检查CAN设备！']]

            if not CAN_option.start(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX, channel):
                # self.showMessageBox(['CAN设备存在问题','CAN通道打开失败，请检查CAN设备！'])
                # self.CANFail()
                return [False, ['CAN设备存在问题', f'CAN通道{channel}打开失败，请检查CAN设备！']]

        return [True, ['', '']]
    def getCANObj(self,receiveID):
        # 接收帧ID
        RECEIVE_ID = receiveID

        # 时间标识
        TIME_STAMP = 0

        # 是否使用时间标识
        TIME_FLAG = 0

        # 接收帧类型
        RECEIVE_SEND_TYPE = 1

        # 是否是远程帧
        REMOTE_FLAG = 0

        # 是否是扩展帧
        EXTERN_FLAG = 0
        DATALEN = 8
        ubyte_array_8 = c_ubyte * 8
        DATA = ubyte_array_8(0, 0, 0, 0, 0, 0, 0, 0)
        ubyte_array_3 = c_ubyte * 3
        RESERVED_3 = ubyte_array_3(0, 0, 0)
        m_can_obj = CAN_option.VCI_CAN_OBJ(RECEIVE_ID, TIME_STAMP,
                                           TIME_FLAG, RECEIVE_SEND_TYPE,
                                           REMOTE_FLAG, EXTERN_FLAG,
                                           DATALEN, DATA, RESERVED_3)
        return m_can_obj

    def  receiveAIData(self,CANAddr_AI,channel):
        can_id = 0x280 + CANAddr_AI
        recv = 0
        time1 = time.time()
        while True:
            if (time.time() - time1)*1000 > self.waiting_time:
                return False,0
            bool_receive,self.m_can_obj = CAN_option.receiveCANbyID(can_id, self.waiting_time,1)
            if bool_receive == 'stopReceive':
                return 'stopReceive', recv
            if bool_receive:
                break
        # self.result_signal.emit(f'self.m_can_obj.Data={hex(self.m_can_obj.Data[0])} {hex(self.m_can_obj.Data[1])} {hex(self.m_can_obj.Data[2])} {hex(self.m_can_obj.Data[3])} {hex(self.m_can_obj.Data[4])} {hex(self.m_can_obj.Data[5])} {hex(self.m_can_obj.Data[6])} {hex(self.m_can_obj.Data[7])}\n\n')
        data = bytes([self.m_can_obj.Data[2*channel], self.m_can_obj.Data[2*channel+1]])
        recv = struct.unpack('<h', data)[0]

        return True,recv

    def clearList(self, array):
        for i in range(len(array)):
            array[i] = 0x00

    # 自动分配节点
    def configCANAddr(self, addr1, addr2, addr3):

        list = [addr1, addr2, addr3]

        for a in list:
            self.m_transmitData = [0xac, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
            self.m_transmitData[2] = a
            boola = CAN_option.transmitCANAddr(0x0, self.m_transmitData)[0]
            time.sleep(0.01)
            if not boola:
                self.result_signal.emit(f'节点{a}分配失败' + self.HORIZONTAL_LINE)
                return False
        self.result_signal.emit(f'所有节点分配成功' + self.HORIZONTAL_LINE)
        return True

    def fillInCPUData(self,station, book, sheet):
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
        self.pass_style = xlwt.XFStyle()
        self.pass_style.borders = pass_borders
        self.pass_style.alignment = pass_alignment
        self.pass_style.font = pass_font
        self.pass_style.pattern = pass_pattern

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
        self.fail_style = xlwt.XFStyle()
        self.fail_style.borders = fail_borders
        self.fail_style.alignment = fail_alignment
        self.fail_style.font = fail_font
        self.fail_style.pattern = fail_pattern

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
        self.warning_style = xlwt.XFStyle()
        self.warning_style.borders = warning_borders
        self.warning_style.alignment = warning_alignment
        self.warning_style.font = warning_font
        self.warning_style.pattern = warning_pattern

        if self.CPU_testItem[0] and self.isPassAppearance:
            sheet.write(self.generalTest_row, 3, '√', self.pass_style)
        elif self.CPU_testItem[0] and not self.isPassAppearance:
            sheet.write(self.generalTest_row, 3, '×', self.fail_style)
            self.errorNum += 1
            self.errorInf += f'\n{self.errorNum})外观存在瑕疵 '
        elif not self.CPU_testItem[0]:
            sheet.write(self.generalTest_row, 3, '未测试', self.warning_style)
        CPU_testName_dict = {1:"型号", 2:"SRAM", 3:"FLASH读写", 4:"拨杆", 5:"MFK按键",
                                   6:"掉电保存", 7:"RTC", 8:"FPGA", 9:"各指示灯", #10:self.inErrorInf, 11:self.outErrorInf,
                                   12:"以太网", 13:"RS-232C", 14:"RS-485", 15:"右扩CAN", 16:"MAC/三码写入", 17:"选项板"}
        for it in range(1,len(self.CPU_passItem)):
            if it <7:
                if self.CPU_testItem[it] and self.CPU_passItem[it]:
                    sheet.write(self.generalTest_row + 2, 3 * it, '√', self.pass_style)
                elif self.CPU_testItem[it] and not self.CPU_passItem[it]:
                    sheet.write(self.generalTest_row + 2, 3 * it, '×', self.fail_style)
                    self.errorInf += f'\n{CPU_testName_dict[it]}测试不通过'
                elif not self.CPU_testItem[it]:
                    sheet.write(self.generalTest_row + 2, 3 * it, '未测试', self.warning_style)
            elif it>=7 and it<13:
                if self.CPU_testItem[it] and self.CPU_passItem[it]:
                    sheet.write(self.generalTest_row + 3, 3 * (it-6), '√', self.pass_style)
                    if it == 10:
                        self.fillCPUInOut(sheet,'in',0)
                    elif it == 11:
                        self.fillCPUInOut(sheet,'out',0)
                elif self.CPU_testItem[it] and not self.CPU_passItem[it]:
                    sheet.write(self.generalTest_row + 3, 3 * (it-6), '×', self.fail_style)
                    if it == 10:
                        self.fillCPUInOut(sheet,'in',1)
                        self.errorInf += f'\n{self.inErrorInf}测试不通过'
                    elif it == 11:
                        self.fillCPUInOut(sheet,'out',1)
                        self.errorInf += f'\n{self.outErrorInf}测试不通过'
                    else:
                        self.errorInf += f'\n{CPU_testName_dict[it]}测试不通过'
                elif not self.CPU_testItem[it]:
                    if it == 10:
                        self.fillCPUInOut(sheet,'in',2)
                    elif it == 11:
                        self.fillCPUInOut(sheet, 'out', 2)
                    sheet.write(self.generalTest_row + 3, 3 * (it-6), '未测试', self.warning_style)
            elif it>=13:
                if self.CPU_testItem[it] and self.CPU_passItem[it]:
                    sheet.write(self.generalTest_row + 4, 3 * (it-12), '√', self.pass_style)
                elif self.CPU_testItem[it] and not self.CPU_passItem[it]:
                    sheet.write(self.generalTest_row + 4, 3 * (it-12), '×', self.fail_style)
                    self.errorInf += f'\n{CPU_testName_dict[it]}测试不通过'
                elif not self.CPU_testItem[it]:
                    sheet.write(self.generalTest_row + 4, 3 * (it-12), '未测试', self.warning_style)

        name_save = ''
        if self.isPassAll and self.testNum == 0:
            name_save = '合格'
            sheet.write(27, 4, '■ 合格', self.pass_style)
            sheet.write(27, 6,
                        '------------------ 全部项目测试通过！！！ ------------------', self.pass_style)
            self.label_signal.emit(['pass','全部通过'])
            self.print_signal.emit([f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}', 'PASS', ''])
        elif self.isPassAll and self.testNum > 0:
            name_save = '部分合格'
            sheet.write(27, 4, '■ 部分合格', self.warning_style)
            sheet.write(27, 6,
                        '------------------ 注意：有部分项目未测试！！！ ------------------', self.warning_style)
            self.label_signal.emit(['testing', '部分通过'])
            self.print_signal.emit([f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}', 'PASS', ''])

        elif not self.isPassAll:
            name_save = '不合格'
            sheet.write(28, 4, '■ 不合格', self.fail_style)
            sheet.write(28, 6, f'不合格原因：{self.errorInf}', self.fail_style)
            self.label_signal.emit(['fail', '未通过'])
            self.print_signal.emit(
                [f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}', 'FAIL', self.errorInf])

        self.saveExcel_signal.emit([book,f'/{name_save}{self.module_type}_{self.module_sn}.xls'])
        # book.save(self.saveDir + f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}.xls')

    def fillCPUInOut(self,sheet,type:str,station:int):
        '''
        :param type: 'in','out'
        :param station: 0:通过，1:不通过，2:未测试
        :return:
        '''
        if station ==0:
            if type == 'in':
                for ch in range(24):
                    if ch < 16:
                        sheet.write(10, ch+3, '√', self.pass_style)
                    else:
                        sheet.write(12, ch-16 + 3, '√', self.pass_style)
            elif type == 'out':
                for ch in range(16):
                    sheet.write(14, ch+3, '√', self.pass_style)
            else:
                self.result_signal.emit('type参数传递错误。\n')
        elif station == 1:
            if type == 'in':
                for ch in range (24):
                    # self.result_signal.emit(f'self.inCh_bin[{ch}}] = {self.inCh_bin[ch]}')
                    if ch < 16:
                        if self.inCh_bin[ch] == '0':
                            sheet.write(10, ch + 3, '×', self.fail_style)
                            self.inErrorInf += f'{ch+1}、'
                        elif self.inCh_bin[ch] == '1':
                            sheet.write(10, ch + 3, '√', self.pass_style)
                    else:
                        if self.inCh_bin[ch] == '0':
                            sheet.write(12, ch-16 + 3, '×', self.fail_style)
                            self.inErrorInf += f'{ch+1}、'
                        elif self.inCh_bin[ch] == '1':
                            sheet.write(12, ch-16 + 3, '√', self.pass_style)
                len_in = len(self.inErrorInf)
                if self.inErrorInf[len_in - 1] == '、':
                    self.inErrorInf = self.inErrorInf[:len_in - 1]
            elif type == 'out':
                for ch in range(16):
                    if self.outCh_bin[ch] == '0':
                        sheet.write(14, ch + 3, '×', self.fail_style)
                        self.outErrorInf += f'{ch+1}、'
                    elif self.outCh_bin[ch] == '1':
                        sheet.write(14, ch + 3, '√', self.pass_style)
                len_out = len(self.outErrorInf)
                if self.outErrorInf[len_out-1] == '、':
                    self.outErrorInf = self.outErrorInf[:len_out-1]
            else:
                self.result_signal.emit('type参数传递错误。\n')
        elif station == 2:
            if type == 'in':
                for ch in range(24):
                    if ch < 16:
                        sheet.write(10, ch+3, '未测试', self.warning_style)
                    else:
                        sheet.write(12, ch-16 + 3, '未测试', self.warning_style)
            elif type == 'out':
                for ch in range(16):
                    sheet.write(14, ch+3, '未测试', self.warning_style)
            else:
                self.result_signal.emit('type参数传递错误。\n')
        else:
            self.result_Signal.emit('station传递错误。\n')