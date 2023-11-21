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
    labe_signal = pyqtSignal(list)
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
    serial_transmitData =[]
    serial_receiveData = []
    # 主副线程弹窗结果
    result_queue = 0

    #各测试项通过标志
    isPassAppearance = False
    isPassTypeCheck = False
    isPassSRAM = False
    isPassFLASH = False
    isPassConfig = False
    isPassFPGA = False
    isPassMFK = False
    isPassRTC = False
    isPassPowerOffSave = False
    isPassLED = False
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
        self.in_Channels = int(inf_CPUlist[1][4])
        self.out_Channels = int(inf_CPUlist[1][4])
        # self.inf_CANIPAdrr = [self.CANAddr1,self.CANAddr2,self.CANAddr3,self.CANAddr4,
        #                                   self.CANAddr5,self.IPAddr]
        # 获取CAN、IP地址
        self.CANAddr1 = inf_CPUlist[2][0]
        self.CANAddr2 = inf_CPUlist[2][1]
        self.CANAddr3 = inf_CPUlist[2][2]
        self.CANAddr4 = inf_CPUlist[2][3]
        self.CANAddr5 = inf_CPUlist[2][4]
        self.IPAddr = inf_CPUlist[2][5]
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
        self.waiting_time = 2000

    def CPUOption(self):
        self.isExcel = True
        #测试是否成功标志
        self.testSign = True
        try:
            ############################################################总线初始化############################################################
            try:
                time_online = time.time()
                while True:
                    QApplication.processEvents()
                    if (time.time() - time_online) * 1000 > 2000:
                        self.pauseOption()
                        if not self.is_running:
                            # 后续测试全部取消
                            self.isTest = False
                            self.CPU_isTest = [False for x in range(self.testNum)]
                            self.testSign = False
                            break
                        self.result_signal.emit(f'错误：总线初始化超时！' + self.HORIZONTAL_LINE)
                        self.messageBox_signal.emit(['错误提示', '总线初始化超时！请检查CAN分析仪或各设备是否正确连接！'])
                        # 后续测试全部取消
                        self.isTest = False
                        self.CPU_isTest = [False for x in range(self.testNum)]
                        self.testSign = False
                        break
                    CANAddr_array = [self.CANAddr1,self.CANAddr2,self.CANAddr3,self.CANAddr4,self.CANAddr5]
                    module_array = [self.module_1, self.module_2, self.module_3, self.module_4, self.module_5]
                    bool_online, eID = otherOption.isModulesOnline(CANAddr_array,module_array, self.waiting_time)
                    if bool_online:
                        self.pauseOption()
                        if not self.is_running:
                            # 后续测试全部取消
                            self.isTest = False
                            self.CPU_isTest = [False for x in range(self.testNum)]
                            self.testSign = False
                            break
                        self.result_signal.emit(f'总线初始化成功！' + self.HORIZONTAL_LINE)
                        break
                    else:
                        self.pauseOption()
                        if not self.is_running:
                            # 后续测试全部取消
                            self.isTest = False
                            self.CPU_isTest = [False for x in range(self.testNum)]
                            self.testSign = False
                            break
                        if eID != 0:
                            self.result_signal.emit(f'错误：未发现{module_array[eID-1]}' + self.HORIZONTAL_LINE)
                        self.result_signal.emit(f'错误：总线初始化失败！再次尝试初始化。' + self.HORIZONTAL_LINE)

                self.result_signal.emit('模块在线检测结束！' + self.HORIZONTAL_LINE)

            except:
                self.pauseOption()
                if not self.is_running:
                    # 后续测试全部取消
                    self.isTest = False
                    self.CPU_isTest = [False for x in range(self.testNum)]
                    self.testSign = False
                self.messageBox_signal.emit(['错误提示', '总线初始化异常，请检查设备！'])
                # 后续测试全部取消
                self.isTest = False
                self.CPU_isTest = [False for x in range(self.testNum)]
                self.testSign = False
                self.result_signal.emit('模块在线检测结束！' + self.HORIZONTAL_LINE)
            ############################################################检测项目############################################################
            for i in range(len(self.CPU_isTest)):
                if not self.CPU_isTest[i]:
                    if i == 0:#外观检测
                        self.CPU_appearanceCheck()
                        if self.isCancelAllTest:
                            break
                    elif i == 1:#型号检查
                        try:
                            self.CPU_typeCheck()
                        except Exception as e:
                            self.showErrorInf()
                        finally:
                            if self.isCancelAllTest:
                                break
                    elif i == 2:#SRAM
                        try:
                            self.CPU_SRAMCheck()
                        except Exception as e:
                            self.showErrorInf()
                        pass
                    elif i == 3:#FLASH
                        pass
                    elif i == 4:#FPGA
                        pass
                    elif i == 5:#拨杆测试
                        pass
                    elif i == 6:#MFK按钮
                        pass
                    elif i == 7:#RTC测试
                        pass
                    elif i == 8:#掉电保存
                        pass
                    elif i == 9:#各指示灯
                        pass
                    elif i == 10:#本体IN
                        pass
                    elif i == 11:#本体OUT
                        pass
                    elif i == 12:#以太网
                        pass
                    elif i == 13:#RS-232C
                        pass
                    elif i == 14:#RS-485
                        pass
                    elif i == 15:#右扩CAN
                        pass
                    elif i == 16:#MA0202
                        pass
                    elif i == 17:#测试报告
                        pass
                    elif i == 18:#固件烧录
                        pass
                    elif i == 19:#MAC/三码写入
                        pass
                    elif i == 20:#U盘读写
                        pass


        except:
            self.messageBox_signal.emit(['错误提示', '测试出现问题，请检查测试程序和测试设备！'])

            self.result_signal.emit('测试出现问题，请检查测试程序和测试设备！' + self.HORIZONTAL_LINE)
            self.allFinished_signal.emit()
            self.pass_signal.emit(False)
    #CPU外观检查
    def CPU_appearanceCheck(self):
        appearanceStart_time = time.time()
        self.item_signal.emit([0, 1, 0, ''])
        self.messageBox_signal.emit(['外观检测', '产品外观是否完好?'])
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
    def CPU_typeCheck(self):
        testStartTime = time.time()
        self.item_signal.emit([1, 1, 0, ''])
        self.clearList(self.m_transmitData)
        self.serial_transmitData = [0xAC,0x06,0x00,0x00,0x0E,0x00]
        #打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=9600, timeout=1)
        loopStartTime = time.time()
        while True:
            if (time.time() - loopStartTime)*1000>2000:
                self.messageBox_signal(['测试警告', '数据接收超时，测试不通过，是否进行后续测试？'])
                self.isPassTypeCheck = False
                reply = self.result_queue.get()
                if reply == QMessageBox.Yes:
                    pass
                elif reply == QMessageBox.No:
                    self.cancelAllTest()
                # 关闭串口
                typeC_serial.close()
                break
            # 发送数据
            typeC_serial.write(bytes(self.serial_transmitData))
            # 等待0.5s后接收数据
            time.sleep(0.5)
            serial_receiveData = []
            # 接收数组数据
            serial_receiveData = typeC_serial.read(30)
            #数据头位置
            startNum = 0
            #确定数据头
            for i in range(len(serial_receiveData)):
                if serial_receiveData[i] == 0xCA:
                    startNum = i
                    isSendAgain = False
                    break
                if i == len(serial_receiveData) - 1 and serial_receiveData[i] != 0xCA:
                    isSendAgain = True
                    break
            if isSendAgain:
                continue
            dataLen = int(serial_receiveData[startNum+1])

            #验证数据接收是否正确
            dataSum = 0
            for i in range(dataLen):
                dataSum+=serial_receiveData[startNum+i]
            ckeckSum = -int(dataSum)
            if ckeckSum == int(serial_receiveData[startNum+dataLen]):
                trueData = serial_receiveData[startNum:startNum+dataLen]
                if trueData[6] == 0x00:#接收成功
                    if trueData[7] == 0x40:
                        self.isPassTypeCheck = True
                        self.changeTabItem(testStartTime,row=1,state=2,result=1)
                        break
                    else:
                        self.isPassTypeCheck = False
                        self.changeTabItem(testStartTime, row=1, state=2, result=2)
                        self.cancelAllTest()
                        self.messageBox_signal(['测试警告', '测试设备型号错误，后续测试中止！'])
                        break
                elif trueData == 0x01:#接收失败
                    continue
            else:
                continue
        # 关闭串口
        typeC_serial.close()

    #SRAM检查
    def CPU_SRAMCheck(self):
        testStartTime = time.time()
        self.item_signal.emit([2, 1, 0, ''])
        self.clearList(self.m_transmitData)
        self.serial_transmitData = [0xAC, 0x06, 0x00, 0x00, 0x0E, 0x00]
        # 打开串口
        typeC_serial = serial.Serial(port=str(self.comboBox_23.currentText()), baudrate=9600, timeout=1)
        # 发送数据
        typeC_serial.write(bytes(self.serial_transmitData))
        # 等待0.5s后接收返回的数据
        time.sleep(0.5)
        # 接收数组数据
        serial_receiveData = typeC_serial.read(6)
        # 关闭串口
        typeC_serial.close()
        if serial_receiveData[5] == '':
            self.isPassTypeCheck = True
            testEndTime = time.time()
            testAllTime = round(testEndTime - testStartTime, 2)
            self.item_signal.emit([2, 2, 1, testAllTime])
        else:
            self.isPassTypeCheck = False
            self.messageBox_signal.emit(['测试警告', '测试设备型号错误，后续测试中止！'])
            testEndTime = time.time()
            testAllTime = round(testEndTime - testStartTime, 2)
            self.item_signal.emit([2, 2, 2, testAllTime])

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
        self.labe_signal.emit(['fail', '测试停止'])

    def showErrorInf(self):
        # 捕获异常并输出详细的错误信息
        self.showInf(f"ErrorInf:\n{traceback.format_exc()}")

