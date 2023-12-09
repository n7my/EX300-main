# -*- coding: utf-8 -*-

import time

from PyQt5.QtCore import QThread, pyqtSignal
from main_logic import *
from CAN_option import *
import otherOption
import logging
import struct

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
        self.CPU_isTest = inf_CPUlist[4][0]
        self.isExcel = inf_CPUlist[4][1]

        #信号等待时间(ms)
        self.waiting_time = 2000

        self.isCancelAllTest = False
        
        self.isPassAll = True
        self.pause_num = 1
        # #初始化表格状态
        # for i in range(len(self.CPU_isTest)):
        #     self.result_signal.emit(f"self.CPU_isTest[{i}]:{self.CPU_isTest[i]}\n")
        #     print(f"self.CPU_isTest[{i}]:{self.CPU_isTest[i]}\n")
        #     if self.CPU_isTest[i]:
        #         self.item_signal.emit([i, 3, 0, ''])
        #     else:
        #         self.item_signal.emit([i, 0, 0, ''])


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
            for i in range(len(self.CPU_isTest)):
                if self.isCancelAllTest:
                    break
                if self.CPU_isTest[i]:
                    if i == 0:#外观检测
                        self.result_signal.emit(f'----------------外观检测----------------')
                        self.CPU_appearanceTest()
                        if self.isCancelAllTest:
                            break
                        self.isPassAll &= self.isPassAppearance
                    elif i == 1:#型号检查
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
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
                            self.showErrorInf('型号检查')
                            self.cancelAllTest()
                        finally:
                            self.isPassAll &= self.isPassTypeTest
                            self.testNum = self.testNum - 1
                            if self.isPassTypeTest:
                                self.result_signal.emit("型号检查通过"+self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("型号检查不通过"+self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 2:#SRAM
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------SRAM测试----------------')
                        try:
                            serial_transmitData = [0xAC, 5, 0x00, 0x01, 0x53]
                            self.CPU_SRAMTest(serial_transmitData=[0xAC, 5, 0x00, 0x01, 0x53,
                                                                 self.getCheckNum(serial_transmitData)])
                        except:
                            self.showErrorInf('SRAM')
                            self.cancelAllTest()
                        finally:
                            self.isPassAll &= self.isPassSRAM
                            self.testNum = self.testNum - 1
                            if self.isPassSRAM:
                                self.result_signal.emit("SRAM测试通过"+self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("SRAM测试不通过"+self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 3:#FLASH
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
                            self.isPassAll &= self.isPassFLASH
                            self.testNum = self.testNum - 1
                            if self.isPassFLASH:
                                self.result_signal.emit("FLASH测试通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("FLASH测试不通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 4:#拨杆测试
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------拨杆测试----------------')
                        try:
                            self.messageBox_signal.emit(["操作提示","请将拨杆拨至STOP位置。"])
                            reply = self.result_queue.get()
                            if reply == QMessageBox.Yes or QMessageBox.No:
                                self.CPU_LeverTest(0,serial_transmitData = [0xAC, 6, 0x00, 0x05, 0x0E, 0x00,
                                                                self.getCheckNum([0xAC, 6, 0x00, 0x05, 0x0E, 0x00])])
                            if not self.isCancelAllTest:
                                self.messageBox_signal.emit(["操作提示", "请将拨杆拨至RUN位置。"])
                                reply = self.result_queue.get()
                                if reply == QMessageBox.Yes or QMessageBox.No:
                                    self.CPU_LeverTest(1,serial_transmitData = [0xAC, 6, 0x00, 0x05, 0x0E, 0x00,
                                                                self.getCheckNum([0xAC, 6, 0x00, 0x05, 0x0E, 0x00])])
                        except:
                            self.showErrorInf('拨杆测试')
                            self.cancelAllTest()
                        finally:
                            self.isPassAll &= self.isPassLever
                            self.testNum = self.testNum - 1
                            if self.isPassLever:
                                self.result_signal.emit("拨杆测试通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("拨杆测试不通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 5:#MFK按钮
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------MFK测试----------------')
                        try:
                            self.CPU_MFKTest(0,serial_transmitData = [0xAC, 6, 0x00, 0x06, 0x0E, 0x00,
                                                                  self.getCheckNum([0xAC, 6, 0x00, 0x06, 0x0E, 0x00])])
                            if not self.isCancelAllTest:
                                self.messageBox_signal.emit(["操作提示", "请长按MKF按钮不要松开。"])
                                reply = self.result_queue.get()
                                if reply == QMessageBox.Yes or QMessageBox.No:
                                    self.CPU_MFKTest(1,serial_transmitData = [0xAC, 6, 0x00, 0x06, 0x0E, 0x00,
                                                                  self.getCheckNum([0xAC, 6, 0x00, 0x06, 0x0E, 0x00])])
                                self.messageBox_signal.emit(["操作提示", "请松开MKF按钮。"])
                                reply = self.result_queue.get()
                                if reply == QMessageBox.Yes or QMessageBox.No:
                                    pass
                        except:
                            self.showErrorInf('MFK按钮')
                            self.cancelAllTest()
                        finally:
                            self.isPassAll &= self.isPassMFK
                            self.testNum = self.testNum - 1
                            if self.isPassMFK:
                                self.result_signal.emit("MFK测试通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("MFK测试不通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 6:#掉电保存
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
                                self.messageBox_signal.emit(['操作提示','请断开设备电源。'])
                                reply= self.result_queue.get()
                                if reply == QMessageBox.Yes:
                                    for t in range(5):
                                        if t == 0:
                                            self.result_signal.emit(f'等待5秒。\n')
                                        self.result_signal.emit(f'剩余等待{5 - t}秒……\n')
                                        time.sleep(1)
                                    self.messageBox_signal.emit(['操作提示', '请接通设备电源。'])
                                    reply = self.result_queue.get()
                                    if reply == QMessageBox.Yes:
                                        for t in range(3):
                                            if t == 0:
                                                self.result_signal.emit(f'等待3秒。\n')
                                            self.result_signal.emit(f'剩余等待{3 - t}秒……\n')
                                            time.sleep(1)
                                        # 数据检查
                                        self.CPU_powerOffSaveTest(serial_transmitData=serial_transmitData1)
                                        if not self.isCancelAllTest:
                                            # 数据清除
                                            self.CPU_powerOffSaveTest(serial_transmitData=serial_transmitData2)
                                    else:
                                        self.messageBox_signal.emit(['测试警告', '未按提示接通设备电源，测试停止。'])
                                        self.isPassPowerOffSave = False
                                        self.cancelAllTest()

                                else:
                                    self.messageBox_signal.emit(['测试警告', '未按提示断开设备电源，测试停止。'])
                                    self.isPassPowerOffSave = False
                                    self.cancelAllTest()

                        except:
                            self.showErrorInf('掉电保存测试')
                            self.cancelAllTest()
                        finally:
                            #若工装的电是通过CPU供电的，断电后需要重新分配节点
                            self.configCANAddr(self.CANAddr1, self.CANAddr2, self.CANAddr3,
                                               self.CANAddr4, self.CANAddr5)
                            self.isPassAll &= self.isPassPowerOffSave
                            self.testNum = self.testNum - 1
                            if self.isPassPowerOffSave:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break


                    elif i == 7:#RTC
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------RTC测试----------------')
                        try:
                            testStartTime = time.time()
                            self.CPU_RTCTest('write')
                            if not self.isCancelAllTest:
                                self.CPU_RTCTest('read')
                        except:
                            self.showErrorInf('RTC测试')
                            self.cancelAllTest()
                        finally:
                            self.testNum = self.testNum - 1
                            if self.isPassRTC:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 8:#FPGA
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------FPGA测试----------------')
                        try:
                            self.messageBox_signal.emit(["操作提示","请插入U盘，并观察U盘指示灯是否点亮？"])
                            reply = self.result_queue.get()
                            if reply == QMessageBox.Yes:
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
                                        self.result_signal.emit('保存固件成功。\n\n')
                                    else:
                                        self.result_signal.emit('保存固件失败。\n\n')
                                    self.CPU_FPGATest(transmitData=transmitData1)
                                    if not self.isCancelAllTest:
                                        if self.isPassFPGA:
                                            self.result_signal.emit('加载固件成功。\n\n')
                                        else:
                                            self.result_signal.emit('加载固件失败。\n\n')
                                        self.CPU_FPGATest(transmitData=transmitData2)
                                        if not self.isCancelAllTest:
                                            if self.isPassFPGA:
                                                self.result_signal.emit('端口配置成功。\n\n')
                                            else:
                                                self.result_signal.emit('端口配置失败。\n\n')
                                        else:
                                            break
                                    else:
                                        break
                                else:
                                    break
                            else:
                                self.messageBox_signal.emit(['测试警告', '无法识别U盘，FPGA无法测试，是否进行后续测试？'])
                                self.isPassFPGA = False
                                reply = self.result_queue.get()
                                if reply == QMessageBox.Yes:
                                    pass
                                else:
                                    self.cancelAllTest()
                        except:
                            self.showErrorInf('FPGA')
                            self.cancelAllTest()
                        finally:
                            self.isPassAll &= self.isPassFPGA
                            self.testNum = self.testNum - 1
                            if self.isPassFPGA:
                                self.result_signal.emit("FPGA测试通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.result_signal.emit("FPGA测试不通过" + self.HORIZONTAL_LINE)
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 9:#各指示灯
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
                            name_LED = ['RUN灯','ERROR灯','BAT_LOW灯','PRG灯','232灯','485灯','HOST灯','所有灯']
                            #进入灯测试模式
                            self.CPU_LEDTest(transmitData=inTest_transmitData,LED_name='')
                            if not self.isCancelAllTest:
                                for l in range(8):
                                    if l != 7:
                                        self.result_signal.emit(f"----------开始测试{name_LED[l]}----------\n")
                                        w_transmitData[5] = l
                                        w_transmitData[7] = self.getCheckNum(w_transmitData[:7])
                                        self.CPU_LEDTest(transmitData=w_transmitData, LED_name=name_LED[l])
                                    else:
                                        #关闭所有LED
                                        w_allLED = [0xAC, 8, 0x00, 0x09, 0x10, 0xAA, 0x00, 0x00, 0x88]
                                        self.CPU_LEDTest(transmitData=w_allLED, LED_name=name_LED[l])
                                    if not self.isCancelAllTest:
                                        if l != 7:
                                            r_transmitData[5] = l
                                            r_transmitData[6] = self.getCheckNum(r_transmitData[:6])
                                            self.CPU_LEDTest(transmitData=r_transmitData, LED_name=name_LED[l])
                                        else:
                                            r_allLED = [0xAC, 6, 0x00, 0x09, 0x0E, 0xAA, 0x8c]
                                            self.CPU_LEDTest(transmitData=r_allLED, LED_name=name_LED[l])

                                    else:
                                        break
                        except:
                            self.showErrorInf('LED测试')
                            self.cancelAllTest()
                        finally:
                            self.testNum = self.testNum - 1
                            # 退出灯测试模式
                            self.CPU_LEDTest(transmitData=outTest_transmitData, LED_name='')
                            if self.isPassLED:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 10:#本体IN
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------本体IN测试----------------')
                        r_transmitData = [0xAC, 7,0x00, 0x0A, 0x0E, 0x00, 0xAA,
                                          self.getCheckNum([0xAC, 7,0x00, 0x0A, 0x0E, 0x00, 0xAA])]
                        try:
                            CAN_init([1])
                            #先控制两块QN0016模块向CPU输入端口发送高电平
                            self.m_transmitData = [0xff,0xff,0x00,0x00,0x00,0x00,0x00,0x00]
                            CAN_option.transmitCAN(0x200 + self.CANAddr2, self.m_transmitData,1)
                            time.sleep(0.1)
                            CAN_option.transmitCAN(0x200 + self.CANAddr3, self.m_transmitData,1)
                            #1.读输入端口
                            self.CPU_InTest(transmitData=r_transmitData)
                            self.m_transmitData = [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
                            CAN_option.transmitCAN(0x200 + self.CANAddr2, self.m_transmitData, 1)
                            time.sleep(0.1)
                            CAN_option.transmitCAN(0x200 + self.CANAddr3, self.m_transmitData, 1)
                            # 2.读输入端口
                            self.CPU_InTest(transmitData=r_transmitData)
                        except:
                            self.showErrorInf('本体IN测试')
                            self.cancelAllTest()
                        finally:
                            self.isPassAll &= self.isPassIn
                            self.testNum = self.testNum - 1
                            if self.isPassIn:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 11:#本体OUT
                        testStartTime = time.time()
                        self.item_signal.emit([i, 1, 0, ''])
                        self.result_signal.emit(f'----------------本体OUT测试----------------')
                        w_transmitData = [0xAC, 10, 0x00, 0x0A, 0x10, 0x01, 0xAA, 2, 0xff, 0xff,
                              self.getCheckNum([0xAC, 10, 0x00, 0x0A, 0x10, 0x01, 0xAA, 2, 0xff, 0xff])]
                        try:
                            CAN_init([1])
                            self.CPU_OutTest(transmitData=w_transmitData)
                            w_transmitData = [0xAC, 10, 0x00, 0x0A, 0x10, 0x01, 0xAA, 2, 0x00, 0x00,
                              self.getCheckNum([0xAC, 10, 0x00, 0x0A, 0x10, 0x01, 0xAA, 2, 0x00, 0x00])]

                            self.CPU_OutTest(transmitData=w_transmitData)
                        except:
                            self.showErrorInf('本体OUT测试')
                            self.cancelAllTest()
                        finally:
                            self.isPassAll &= self.isPassOut
                            self.testNum = self.testNum - 1
                            if self.isPassOut:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 12:#以太网
                        self.item_signal.emit([i, 1, 0, ''])
                        testStartTime = time.time()
                        self.result_signal.emit(f'----------------以太网测试----------------')
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
                            self.testNum = self.testNum - 1
                            if self.isPassETH:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 13:#RS-232C
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
                            self.showErrorInf('本体232测试')
                            self.cancelAllTest()
                        finally:
                            self.testNum = self.testNum - 1
                            self.isPassAll &= self.isPass232
                            if self.isPass232:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 14:#RS-485
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
                            self.showErrorInf('本体485测试')
                            self.cancelAllTest()
                        finally:
                            self.testNum = self.testNum - 1
                            self.isPassAll &= self.isPass485
                            if self.isPass485:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 15:#右扩CAN
                        self.item_signal.emit([i, 1, 0, ''])
                        testStartTime = time.time()

                        transmitData_reset = [0xAC,7, 0x00, 0x0E, 0x10, 0x00, 0x00,
                                              self.getCheckNum([0xAC,7, 0x00, 0x0E, 0x10, 0x00, 0x00])]
                        transmitData_config = [0xAC, 7, 0x00, 0x0E, 0x10, 0x01, 0x01,
                                               self.getCheckNum([0xAC,7, 0x00, 0x0E, 0x10, 0x01, 0x01])]
                        try:
                            self.CPU_rightCANTest(transmitData=transmitData_reset)
                            if not self.isCancelAllTest:
                                transmitData_reset[6] = 0x01
                                transmitData_reset[7] = self.getCheckNum(transmitData_reset[:7])
                                self.CPU_rightCANTest(transmitData=transmitData_reset)
                                if not self.isCancelAllTest:
                                    self.CPU_rightCANTest(transmitData=transmitData_config)
                                    if not self.isCancelAllTest:
                                        transmitData_config[6] = 0x00
                                        transmitData_config[7] = self.getCheckNum(transmitData_config[:7])
                                        self.CPU_rightCANTest(transmitData=transmitData_config)
                                        if not self.isCancelAllTest:
                                            try:
                                                # 连接CAN设备CANalyst-II，并初始化两个CAN通道
                                                self.CAN_init([0])
                                            except Exception as e:
                                                self.result_signal.emit(
                                                    f'CAN_init error:{e}\nCAN设备通道0启动失败，后续测试全部取消。\n')
                                                self.messageBox_signal.emit(['警告',
                                                                             f'CAN_init error:{e}\nCAN设备通道0启动失败，'
                                                                             f'后续测试全部取消。\n'])
                                                reply = self.result_queue.get()
                                                if reply == QMessageBox.Yes:
                                                    self.cancelAllTest()
                                                else:
                                                    self.cancelAllTest()
                                                break
                                            for loopNum in range(32):
                                                rightCAN_transmitData = [(x+8*loopNum) for x in range(8)]
                                                isBool, whatEver = CAN_option.transmitCAN(0x701,rightCAN_transmitData,0)
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
                                                        m_can_obj = self.getCANObj(receiveID= 0x601)
                                                        bool_receive, m_can_obj = CAN_option.receiveCANbyID(0x601,
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
                                                    ['测试警告', '右扩CAN测试不通过，是否进行后续测试？'])
                                                reply = self.result_queue.get()
                                                if reply == QMessageBox.Yes:
                                                    pass
                                                else:
                                                    self.cancelAllTest()
                                        else:
                                            break
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
                            self.isPassAll &= self.isPassRightCAN
                            self.testNum = self.testNum - 1
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
                            self.showErrorInf('三码与MAC地址写入测试')
                            self.cancelAllTest()
                        finally:
                            self.testNum = self.testNum - 1
                            if self.isPassConfig:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
                                self.changeTabItem(testStartTime, row=i, state=2, result=2)
                            if self.isCancelAllTest:
                                break
                    elif i == 17:#MA0202
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
                                                    if reply == QMessageBox.Yes:
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
                                                            CAN_option.transmitCAN((0x600 + self.CANAddr5),
                                                                                   AQ_transmit,1)
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
                                                    if reply == QMessageBox.Yes:
                                                        self.cancelAllTest()
                                                    else:
                                                        self.cancelAllTest()
                                                    break
                                            else:
                                                break
                                        else:
                                            break
                                    else:
                                        break
                                else:
                                    break
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
                                                    if reply == QMessageBox.Yes:
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
                                                if reply == QMessageBox.Yes:
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
                            self.showErrorInf('选项板测试')
                            self.cancelAllTest()
                        finally:
                            if not self.is_running:
                                self.isPassOp &= False

                            self.isPassAll &= self.isPassOp
                            self.testNum = self.testNum - 1
                            if self.isPassOp:
                                self.changeTabItem(testStartTime, row=i, state=2, result=1)
                            else:
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
                if reply == QMessageBox.Yes or reply == QMessageBox.No:
                    self.messageBox_signal.emit(['操作提示', f'请将拨杆拨到STOP位置。\n'])
            else:
                self.messageBox_signal.emit(['操作提示', f'测试被中止！\n'])
                self.label_signal.emit(['fail', '测试停止'])

        except:
            self.isPassAll &= False
            self.messageBox_signal.emit(['错误提示', f'测试出现问题，请检查测试程序和测试设备！\n'
                                                     f'"ErrorInf:\n{traceback.format_exc()}"'])
            self.showErrorInf('测试')
        finally:
            # if self.isExcel:
            #     self.result_signal.emit('开始生成校准校验表…………' + self.HORIZONTAL_LINE)
            #     code_array = [self.module_pn, self.module_sn, self.module_rev,self.module_MAC]
            #     station_array = [self.isAOPassTest, self.isAOTestVol, self.isAOTestCur]
            #
            #     excel_bool, book, sheet, self.CPU_row = otherOption.generateExcel(code_array,
            #                                                                      station_array, 0, 'CPU')
            #     if not excel_bool:
            #         self.result_signal.emit('校准校验表生成出错！请检查代码！' + self.HORIZONTAL_LINE)
            #     else:
            #         self.fillInAOData(self.isAOPassTest, book, sheet)
            #         self.result_signal.emit('生成校准校验表成功' + self.HORIZONTAL_LINE)
            #
            #     if not self.isAOVolPass or not self.isAOCurPass:
            #         self.result_signal.emit(f'不通过原因：\n{self.errorInf}' + self.HORIZONTAL_LINE)
            #
            # elif not self.isExcel:
            #     self.result_signal.emit('测试停止，未生成校准校验表！' + self.HORIZONTAL_LINE)

            self.allFinished_signal.emit()
            self.pass_signal.emit(False)
    #CPU外观检查
    def CPU_appearanceTest(self):
        appearanceStart_time = time.time()
        self.item_signal.emit([0, 1, 0, ''])
        self.messageBox_signal.emit(['外观检测', '请检查：\n（1）外壳字体是否清晰?\n（2）型号是否正确？\n（3）外壳是否完好？'])
        reply = self.result_queue.get()
        if reply == QMessageBox.Yes:
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
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                                                    f'Failed to open serial port: {e}'])
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
            if reply == QMessageBox.Yes or reply == QMessageBox.No:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    #SRAM测试
    def CPU_SRAMTest(self,serial_transmitData:list):

        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                                                    f'Failed to open serial port: {e}'])
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
            self.messageBox_signal.emit(['测试警告', 'SRAM测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            else:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    #FLASH测试
    def CPU_FLASHTest(self,serial_transmitData:list): 
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                                                    f'Failed to open serial port: {e}'])
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
            self.messageBox_signal.emit(['测试警告', 'FLASH测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            else:
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
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                                                    f'Failed to open serial port: {e}'])
        loopStartTime = time.time()

        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassFPGA &= False
                isThisPass = False
                break
            # 发送数据
            typeC_serial.write(bytes(transmitData))
            name_dict ={0x00:'保存固件',0x01:'加载固件',0x02:'端口配平'}
            self.result_signal.emit(f'等待<{name_dict[transmitData[5]]}>自测。\n\n')
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
            self.messageBox_signal.emit(['测试警告', 'FPGA测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            else:
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
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                                                    f'Failed to open serial port: {e}'])
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
            self.messageBox_signal.emit(['测试警告', '拨杆测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            else:
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
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                                                    f'Failed to open serial port: {e}'])
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
            self.messageBox_signal.emit(['测试警告', 'MFK测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            else:
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
            self.messageBox_signal.emit(['错误警告', f'Failed to open serial port: {e}'])
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
            self.messageBox_signal.emit(['测试警告', '掉电保存测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            else:
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

        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                                                    f'Failed to open serial port: {e}'])
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassRTC = False
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
            self.messageBox_signal.emit(['测试警告', 'RTC测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            else:
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
                self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                                                        f'Failed to open serial port: {e}'])
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
                        self.messageBox_signal.emit(['操作提示',f'{LED_name}是否熄灭？'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.Yes:
                            self.isPassLED &= True
                            isThisPass = True
                        else:
                            self.isPassLED &= False
                            isThisPass = False
                        break
                    elif trueData[7] == 0x01:#灯点亮
                        self.messageBox_signal.emit(['操作提示',f'{LED_name}是否点亮？'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.Yes:
                            self.isPassLED &= True
                            isThisPass = True
                        else:
                            self.isPassLED &= False
                            isThisPass = False
                        break
                elif dataLen == 9:#读所有灯状态
                    if trueData[7] == 0x00 and trueData[8] == 0x00:
                        self.messageBox_signal.emit(['操作提示', f'{LED_name}是否熄灭？'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.Yes:
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
                self.messageBox_signal.emit(['测试警告', 'LED测试不通过，是否进行后续测试？'])
                reply = self.result_queue.get()
                if reply == QMessageBox.Yes:
                    pass
                else:
                    self.cancelAllTest()
            # 关闭串口
            typeC_serial.close()

    def CPU_InTest(self, transmitData:list):
        isThisPass = True
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                                                    f'Failed to open serial port: {e}'])
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
                    if trueData[9] == 0xff and trueData[10] == 0xff and trueData[11] == 0xff:
                        self.messageBox_signal.emit(['操作提示','输入通道指示灯是否全部点亮？'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.Yes:
                            self.isPassIn &= True
                            isThisPass = True
                        else:
                            self.isPassIn &= False
                            isThisPass = False
                        break
                    elif trueData[9] == 0x00 and trueData[10] == 0x00 and trueData[11] == 0x00:
                        self.messageBox_signal.emit(['操作提示','输入通道指示灯是否全部熄灭？'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.Yes:
                            self.isPassIn &= True
                            isThisPass = True
                        else:
                            self.isPassIn &= False
                            isThisPass = False
                        break
                    else:
                        self.isPassIn &= False
                        isThisPass = False
                        break
                elif trueData[7] == 0x01:  # 读所有输入通道失败
                    self.isPassIn &= False
                    isThisPass = False
                    break

        if not isThisPass:
            self.messageBox_signal.emit(['测试警告', '本体IN测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            else:
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
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                                                    f'Failed to open serial port: {e}'])
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
                    elif (hex(self.m_receiveData[0]) == '0xff' and hex(self.m_receiveData[1]) == '0xff') or\
                        (hex(self.m_receiveData[0]) == '0x0' and hex(self.m_receiveData[1]) == '0x0'):
                        if hex(self.m_receiveData[0]) == '0xff':
                            self.messageBox_signal.emit(['操作提示','ET1600的通道指示灯是否全部点亮？'])
                        elif hex(self.m_receiveData[0]) == '0x0':
                            self.messageBox_signal.emit(['操作提示', 'ET1600的通道指示灯是否全部熄灭？'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.Yes:
                            self.isPassOut &= True
                            isThisPass = True
                        else:
                            self.isPassOut &= False
                            isThisPass = False
                        break
                elif trueData[7] == 0x01:  # 读所有输出通道失败
                    self.isPassOut &= False
                    isThisPass = False
                    break
            else:
                continue

        if not isThisPass:
            self.messageBox_signal.emit(['测试警告', '本体OUT测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            else:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    # 本体ETH测试
    def CPU_ETHTest(self, transmitData: list,isPassSign:bool):
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                                                    f'Failed to open serial port: {e}'])
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassETH = False
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
            self.messageBox_signal.emit(['测试警告', '本体ETH IP测试不通过，是否进行后续测试？'])
            self.isPassETH = False
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            else:
                self.cancelAllTest()
        elif (not self.isPassETH_UDP and self.isPassETH_IP) or (not self.isPassETH_UDP and not self.isPassETH_IP):
            self.isPassETH = False
            self.messageBox_signal.emit(['测试警告', '本体ETH UDP测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            else:
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
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                                                    f'Failed to open serial port: {e}'])
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
            self.messageBox_signal.emit(['测试警告', '本体232测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            else:
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
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{m_port}打开失败，请检查该串口是否被占用。\n'
                                                    f'Failed to open serial port: {e}'])
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
                self.messageBox_signal.emit(['操作警告','未接收到信号，请检查232（485）接线是否断开。'])
                reply = self.result_queue.get()
                if reply == QMessageBox.Yes:
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
            self.messageBox_signal.emit(['测试警告', '本体232测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            else:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()
        return isReceiveTrueData




    # 本体485测试
    def CPU_485Test(self, transmitData: list,baudRate:list):
        isThisPass = True
        try:
            # 打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(
                ['错误警告', f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                             f'Failed to open serial port: {e}'])
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
            self.messageBox_signal.emit(['测试警告', '本体485测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            else:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    # 本体右扩CAN测试
    def CPU_rightCANTest(self, transmitData: list):
        isThisPass = True
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                                                    f'Failed to open serial port: {e}'])
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassRightCAN &= False
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
                self.isPassRightCAN &= False
                isThisPass = False
                break
            elif dataLen == 7:  # 写
                if trueData[6] == 0x00:  # 写成功
                    if transmitData[5] == 0x00 and transmitData[6] == 0x00:#复位灯熄灭
                        self.messageBox_signal.emit(['操作提示','右扩CAN的RESET灯是否熄灭?'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.Yes:
                            self.isPassRightCAN &= True
                            isThisPass = True
                        else:
                            self.isPassRightCAN &= False
                            isThisPass = False
                    elif transmitData[5] == 0x00 and transmitData[6] == 0x01:#复位灯亮
                        self.messageBox_signal.emit(['操作提示','右扩CAN的RESET灯是否点亮?'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.Yes:
                            self.isPassRightCAN &= True
                            isThisPass = True
                        else:
                            self.isPassRightCAN &=False
                            isThisPass = False
                    elif transmitData[5] == 0x01 and transmitData[6] == 0x01:#使能灯灭
                        self.messageBox_signal.emit(['操作提示','右扩CAN的使能灯是否熄灭?'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.Yes:
                            self.isPassRightCAN &= True
                            isThisPass = True
                        else:
                            self.isPassRightCAN &=False
                            isThisPass = False
                    elif transmitData[5] == 0x01 and transmitData[6] == 0x00:#使能灯亮
                        self.messageBox_signal.emit(['操作提示','右扩CAN的使能灯是否点亮?'])
                        reply = self.result_queue.get()
                        if reply == QMessageBox.Yes:
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
            self.messageBox_signal.emit(['测试警告', '本体右扩CAN测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            else:
                self.cancelAllTest()
        # 关闭串口
        typeC_serial.close()

    #MAC地址和3码
    def CPU_configTest(self, transmitData: list, type:str):
        try:
            #打开串口
            typeC_serial = serial.Serial(port=str(self.serialPort_typeC), baudrate=1000000, timeout=1)
        except serial.SerialException as e:
            self.messageBox_signal.emit(['错误警告',f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                                                    f'Failed to open serial port: {e}'])
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime) * 1000 > self.waiting_time:
                self.isPassConfig = False
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
            self.messageBox_signal.emit(['测试警告', '本体Config测试不通过，是否进行后续测试？'])
            reply = self.result_queue.get()
            if reply == QMessageBox.Yes:
                pass
            else:
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
                    ['错误警告', f'串口{str(self.serialPort_typeC)}打开失败，请检查该串口是否被占用。\n'
                                 f'Failed to open serial port: {e}'])
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
                            if reply == QMessageBox.Yes:
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
                            if reply == QMessageBox.Yes:
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
                            if reply == QMessageBox.Yes:
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
                            if reply == QMessageBox.Yes:
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
                self.messageBox_signal.emit(['测试警告', '选项板测试不通过，是否进行后续测试？'])
                reply = self.result_queue.get()
                if reply == QMessageBox.Yes:
                    pass
                else:
                    self.cancelAllTest()
            # 关闭串口
            typeC_serial.close()
            return opType

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
        self.result_signal.emit(f'receiveData:{hex_trueData}\n')
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
        self.result_signal.emit(f"{Inf}程序出现问题\n")
        # 捕获异常并输出详细的错误信息
        self.result_signal.emit(f"ErrorInf:\n{traceback.format_exc()}\n")
        self.messageBox_signal.emit(['错误警告',(f"ErrorInf:\n{traceback.format_exc()}\n")])

    #CAN设备初始化
    def CAN_init(self,CAN_channel:list):
        CAN_option.close(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX)
        time.sleep(0.1)
        QApplication.processEvents()

        for channel in CAN_channel:
            #通道0
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
    def configCANAddr(self, addr1, addr2, addr3, addr4, addr5):

        list = [addr1, addr2, addr3, addr4, addr5]

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