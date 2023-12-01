import os
import sys
import threading
import win32api
import win32print
import CAN_option
# from main_ui import Ui_Form
from ctypes import *
import time
from queue import Queue
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from CAN_option import *
from abc import ABCMeta,abstractmethod
# from workThread import WorkThread
from new1440_all import Ui_Form

import psutil
import xlwt
import signal
from PyQt5 import QtCore, QtGui, QtWidgets

import threading
import queue
import serial.tools.list_ports
import traceback


q = queue.Queue()
event_canInit = threading.Event()
isMainRunning = True
class Ui_Control(QMainWindow,Ui_Form):
    #                -/+10V       0-10V       0-5V        1-5V       4-20mA      0-20mA      -/+5V
    AORangeArray = [0x0b, 0x00, 0xe8, 0x03, 0xe9, 0x03, 0xea, 0x03, 0x15, 0x00, 0xeb, 0x03, 0xec, 0x03]
    AIRangeArray = [0x29, 0x00, 0x2a, 0x00, 0x10, 0x27, 0x11, 0x27, 0x33, 0x00, 0x34, 0x00, 0x12, 0x27]

    currentTheory = [0x6c00, (int)(0x6c00 * 0.75), (int)(0x6c00 * 0.5), (int)(0x6c00 * 0.25), 0]  # 0x6c00 =>27648
    arrayCur = ["20mA测试", "15mA测试", "10mA测试", "5mA测试", "0mA测试"]
    highCurrent = 59849
    lowCurrent = 32768

    voltageTheory = [0x6c00, 0x3600, 0x00, 0xca00, 0x9400]  # 0x6c00 =>27648
    arrayVol = ["10V测试", "5V测试", "0V测试", "-5V测试", "-10V测试"]

    highVoltage = 59849
    lowVoltage = 5687

    HORIZONTAL_LINE = "\n------------------------------------------------------------------------------------------------------------\n\n"

    module_type = ''
    module_pn = ''
    module_sn = ''
    module_rev = ''
    # AI模块CAN地址
    CANAddr_AI = 1
    # AO模块CAN地址
    CANAddr_AO = 2
    # DI模块CAN地址
    CANAddr_DI = 1
    # DO模块CAN地址
    CANAddr_DO = 2
    # 继电器CAN地址
    CANAddr_relay = 3
    # 波特率
    baud_rate = 1000
    # 信号等待时间
    waiting_time = 5000
    # 标定接收信号次数
    receive_num = 8
    #循环次数
    loop_num = 1
    pause_num = 1

    #生成表格的相关标识
    #DI测试是否通过
    isDIPassTest = True
    #DO测试是否通过
    isDOPassTest = True
    #DO通道数据初始化
    DO_channelData = 0
    #DI通道数据初始化
    DI_channelData = 0
    # DO通道错误数据记录
    DODataCheck = [True for i in range(32)]
    # DI通道错误数据记录
    DIDataCheck = [True for i in range(32)]



    #是否标定
    isCalibrate = False
    # 是否标定电压
    isCalibrateVol = False
    # 是否标定电流
    isCalibrateCur = False

    # 是否检测
    isCalibrate = False
    # 是否检测AI电压
    isAITestVol = False
    # 是否检测AI电流
    isAITestCur = False
    # 是否检测AO电压
    isAOTestVol = False
    # 是否检测AO电流
    isAOTestCur = False
    # 是否检测CAN_Run+CAN_Error
    isTestCANRunErr = False
    # 是否检测Run+Error
    isTestRunErr = False
    # AI测试是否通过
    isAIPassTest = True
    isAIVolPass = True
    isAICurPass = True
    # AO测试是否通过
    isAOPassTest = True
    isAOVolPass = True
    isAOCurPass = True

    isLEDRunOK = True
    isLEDErrOK = True
    isLEDPass = True

    #CAN帧结构体
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

    #测试总数
    testNum = 0
    #测试类型
    testType = {}
    #模块信息
    inf = {'AO2':'待检AO模块', 'AI1':'配套AI模块', 'AI2':'待检AI模块', 'AO1':'配套AO模块',
           'DO2':'待检DO模块', 'DI1':'配套DI模块', 'DI2':'待检DI模块', 'DO1':'配套DO模块'}
    module_1 = ''
    module_2 = ''
    #通道数
    m_Channels = 0
    AI_Channels = 0
    AO_Channels = 0
    #发送的数据
    ubyte_array_transmit = c_ubyte * 8
    m_transmitData = ubyte_array_transmit(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
    #接收的数据
    ubyte_array_receive = c_ubyte * 8
    m_receiveData = ubyte_array_receive(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
    CAN_errorLED = True
    CAN_runLED = True
    errorLED = True
    runLED = True
    errorNum = 0
    errorInf = ''
    isAllScreen = False
    volReceValue = [0, 0, 0, 0, 0]
    curReceValue = [0, 0, 0, 0, 0]
    volPrecision = [0, 0, 0, 0, 0]
    curPrecision = [0, 0, 0, 0, 0]
    chPrecision = [0, 0, 0, 0]

    subRunFlag = True
    # 创建队列用于主线程和子线程之间的通信
    result_queue = Queue()

    #停止CAN接收信号
    stop_signal = pyqtSignal(bool)

    #3码转换后的ASCII码
    asciiCode_pn = []
    asciiCode_sn = []
    asciiCode_rev = []
    testFlag = ''

    config_param = { }

    current_dir = os.getcwd().replace('\\','/')+"/_internal"
    # current_dir = os.getcwd().replace('\\', '/')
    def __init__(self,parent = None):
        super(Ui_Control,self).__init__(parent)
        self.setupUi(self)
        # self.pushButton_13.setVisible(False)
        self.label_11.setPixmap(QtGui.QPixmap(f"{current_dir}/beast5.png"))
        # CPU页面参数配置
        self.CPU_lineEdit_array = [self.lineEdit_33, self.lineEdit_34, self.lineEdit_35, self.lineEdit_36,
                                   self.lineEdit_37, self.lineEdit_38]
        self.CPU_lineEditName_array = ["CPU_IP", "工装1", "工装2", "工装3", "工装4", "工装5"]
        self.CPU_checkBox_array = [self.checkBox_50, self.checkBox_51, self.checkBox_52, self.checkBox_53,
                                    self.checkBox_55, self.checkBox_56, self.checkBox_57,
                                   self.checkBox_58, self.checkBox_59, self.checkBox_60, self.checkBox_61,
                                   self.checkBox_62, self.checkBox_63, self.checkBox_64, self.checkBox_65,
                                   self.checkBox_66,self.checkBox_54, self.checkBox_67, self.checkBox_68,
                                   self.checkBox_71,self.checkBox_73]
        self.CPU_checkBoxName_array = ["外观检测", "型号检查", "SRAM", "FLASH", "FPGA", "拨杆测试",
                                       "MFK按键",
                                       "掉电保存", "RTC测试", "各指示灯", "本体IN", "本体OUT", "以太网", "RS-232C",
                                       "RS-485",
                                       "右扩CAN", "MAC/三码写入",  "MA0202", "测试报告", "修改参数","全选"]

        self.setWindowFlags(Qt.FramelessWindowHint)  # 无边框窗口
        self.label_6.mousePressEvent = self.label_mousePressEvent
        self.label_6.mouseMoveEvent = self.label_mouseMoveEvent
        self.label.setStyleSheet(self.testState_qss['stop'])
        self.label.setText('UNTESTED')
        self.label_41.setAlignment(QtCore.Qt.AlignLeft|Qt.AlignVCenter)
        #屏蔽pushbutton_9
        self.pushButton_9.setVisible(False)
        self.getSerialInf(0)
        # 读页面配置
        # self.loadConfig()必须放在类似comboBox.currentIndexChanged.connect(self.saveConfig)的代码之前
        # 某则一修改参数就会触发saveConfig，导致还没修改的参数被默认参数覆盖。
        self.loadConfig()
        if self.checkBox_73.isChecked():
            self.checkBox_73.setText("取消全选")
        else:
            self.checkBox_73.setText("全选")

        #显示当前时间
        self.update_time()
        # self.lineEdit_PN = 'PRDr9HPA06Mz-00'
        # self.lineEdit_SN = 'S1223-001083'
        # self.lineEdit_REV = '06'
        self.tabWidget.setTabPosition(0)
        self.tabWidget.setGeometry(QtCore.QRect(750, 314, 672, 530))
        self.tab_bar = self.tabWidget.tabBar()
        # self.textBrowser_5.setStyleSheet("""
        #                     QTextBrowser::pane {
        #                         border: 1px solid #11826C;
        #                         padding: 2px;
        #                     }
        #                 """)
        # for tW in [self.tableWidget_AI,self.tableWidget_AO,self.tableWidget_DIDO]:
        #     tW.setStyleSheet("""
        #                         QtWidgets::pane {
        #                             border: 1px solid #11826C;
        #                             padding: 2px;
        #                         }
        #                     """)




        for tW in [self.tableWidget_AI, self.tableWidget_AO, self.tableWidget_DIDO,self.tableWidget_CPU]:
            tW.setEditTriggers(QAbstractItemView.NoEditTriggers)

        #批量设置lineEdit只读
        lE_arr = [self.lineEdit,self.lineEdit_3,self.lineEdit_5,self.lineEdit_20,self.lineEdit_21,self.lineEdit_22,
                  self.lineEdit_45,self.lineEdit_46,self.lineEdit_47,self.lineEdit_30,self.lineEdit_31,self.lineEdit_32
                  ,self.lineEdit_33]
        for lE in lE_arr:
            lE.setReadOnly(True)

        # 设置选中和未选中状态下的样式
        self.selected_color = QColor("#11826C")
        self.unselected_color = QColor("white")
        self.tabWidget.setStyleSheet("""
                    QTabWidget::pane {
                        border: 1px solid #11826C;
                        padding: 2px;
                    }
                """)
        self.tab_bar.setStyleSheet(f"""
                    QTabBar::tab:selected {{
                        background-color: #F5BD3D;
                        color: #212B36;
                        border-top-left-radius: 0px ;
                        border-top-right-radius: 0px;
                        padding: 8px;
                        width: 120px; 
                        height: 17px;
                        
                    }}
                    QTabBar::tab:!selected {{
                        background-color: #11826C;
                        color: white;
                        border-top-left-radius: 0px;
                        border-top-right-radius: 0px;
                        padding: 8px;
                        width: 120px; 
                        height: 13px;
                        margin-top: 5px
                    }}
                    QTabBar::tab:first {{
                        margin-left: 0;
                    }}
                    QTabBar::tab:last {{
                        margin-right: 0;
                    }}
                    QTabBar::tab {{
                        padding: 5px;
                        margin-right: 2px;
                        margin-bottom: 0px;
                    }}
                    QTabBar::tab:hover {{
                        background-color: #F5BD3D;
                        color: #212B36;
                        border-top-left-radius: 0px ;
                        border-top-right-radius: 0px;
                        padding: 8px;
                        width: 120px; 
                        height: 17px;
                    }}
                 """)

        # 设置字体为黑体加粗
        self.tabBarFont = QFont()
        self.tabBarFont.setFamily("Arial")
        self.tabBarFont.setBold(True)
        self.tabBarFont.setPointSize(12)
        self.tab_bar.setFont(self.tabBarFont)

        # 设置按钮悬停时的样式
        self.hover_color = QColor("#11826C").lighter(150)
        listPushButton = [self.pushButton_5,self.pushButton_6,self.pushButton_10,self.pushButton_renewSerial]
        for pB in listPushButton:
            pB.setStyleSheet(f"""
                        QPushButton {{
                            background-color: #11826C;
                            color:white
                        }}
                        QPushButton:hover {{
                            background-color: {self.hover_color.name()};
                            color:#212B36
                        }}
                    """)

        self.dragging = False
        self.offset = QPoint()


        self.lineEdit_PN.setReadOnly(True)
        self.pushButton_8.clicked.connect(self.killUi)
        self.pushButton_11.clicked.connect(self.showMinimized)
        self.pushButton_11.clicked.connect(self.saveConfig)
        self.tableWidget_array=[self.tableWidget_DIDO,self.tableWidget_AI,
                                self.tableWidget_AO,self.tableWidget_CPU]
        for tW in self.tableWidget_array:
            tW.setStyleSheet("background-color: rgb(255, 255, 255);")
        # self.tableWidget_AI.setStyleSheet("background-color: rgb(255, 255, 255);")
        # self.tableWidget_AO.setStyleSheet("background-color: rgb(255, 255, 255);")
        # self.tableWidget_CPU.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.pushButton_4.clicked.connect(self.start_button)
        self.pushButton_3.clicked.connect(self.stop_button)


        self.pushButton_pause.clicked.connect(self.pause_button)

        self.pushButton_resume.clicked.connect(self.resume_button)

        self.pushButton.setEnabled(True)
        self.pushButton.setStyleSheet(self.topButton_qss['on'])

        self.pushButton_2.setEnabled(True)
        self.pushButton_2.setStyleSheet(self.topButton_qss['on'])

        listTopButton = [self.pushButton, self.pushButton_2]
        for tB in listTopButton:
            tB.setStyleSheet(f"""
                                    QPushButton {{
                                        background-color: #11826C;
                                        color:white;
                                        font: bold 25pt \"宋体\"
                                    }}
                                    QPushButton:hover {{
                                        background-color: {self.hover_color.name()};
                                        color:#212B36;
                                        font: bold 25pt \"宋体\"
                                    }}
                                """)

        self.pushButton_3.setEnabled(False)
        self.pushButton_3.setStyleSheet(self.topButton_qss['off'])

        self.pushButton_4.setEnabled(False)
        self.pushButton_4.setStyleSheet(self.topButton_qss['off'])
        # # 设置按钮悬停时的样式
        # self.hover_pause = QColor("#ffff00").lighter(50)
        self.pushButton_pause.setStyleSheet(self.topButton_qss['pause_on'])
        self.pushButton_pause.setStyleSheet(f"""
                                QPushButton {{
                                    {self.topButton_qss['pause_on']}
                                }}
                                QPushButton:hover {{
                                    background-color: rgb(255, 255, 0);
                                    color:rgb(0, 0, 0)
                                }}
                            """)
        self.pushButton_pause.setEnabled(False)
        self.pushButton_pause.setVisible(False)
        # 设置按钮悬停时的样式
        self.hover_resume = QColor("#00ff00").lighter(50)
        self.pushButton_resume.setStyleSheet(self.topButton_qss['start_on'])
        self.pushButton_resume.setStyleSheet(f"""
                        QPushButton {{
                            {self.topButton_qss['start_on']}
                        }}
                        QPushButton:hover {{
                            background-color: rgb(0, 255, 0);
                            color:rgb(0, 0, 0)
                        }}
                    """)
        self.pushButton_resume.setEnabled(False)
        self.pushButton_resume.setVisible(False)
        for tW in self.tableWidget_array:
            tW.horizontalHeader().setStyleSheet("background-color: #11826C;")
        # self.tableWidget_AO.horizontalHeader().setStyleSheet("background-color: #11826C;")
        # self.tableWidget_DIDO.horizontalHeader().setStyleSheet("background-color: #11826C;")
        # self.tableWidget_CPU.horizontalHeader().setStyleSheet("background-color: #11826C;")

        self.checkBox_5.setEnabled(not self.radioButton.isChecked())
        self.checkBox_6.setEnabled(not self.radioButton.isChecked())
        self.checkBox_7.setEnabled(self.radioButton_5.isChecked())
        self.checkBox_8.setEnabled(self.radioButton_5.isChecked())
        self.checkBox_9.setEnabled(self.radioButton_5.isChecked())
        self.checkBox_10.setEnabled(self.radioButton_5.isChecked())

        self.checkBox_29.setEnabled(self.radioButton_17.isChecked())
        self.checkBox_30.setEnabled(self.radioButton_17.isChecked())
        self.checkBox_31.setEnabled(self.radioButton_17.isChecked())
        self.checkBox_32.setEnabled(self.radioButton_17.isChecked())

        self.checkBox_35.setEnabled(not self.radioButton_18.isChecked())
        self.checkBox_36.setEnabled(not self.radioButton_18.isChecked())

        self.pushButton_5.clicked.connect(self.changeSaveDir)
        self.pushButton_6.clicked.connect(self.openSaveDir)

        # self.pushButton.clicked.connect(self.stopAll)
        self.tabIndex = self.tabWidget.currentIndex()
        if self.tabIndex == 0:
            self.tableWidget_DIDO.setVisible(True)
            self.tableWidget_AI.setVisible(False)
            self.tableWidget_AO.setVisible(False)
            self.tableWidget_CPU.setVisible(False)
            self.label_7.setText("DI/DO")
            if self.comboBox.currentIndex()==1:
                self.lineEdit_PN.setText('P0000010391877')
            elif self.comboBox.currentIndex()==4:
                self.lineEdit_PN.setText('P0000010392086')
            elif self.comboBox.currentIndex()==5:
                self.lineEdit_PN.setText('P0000010392121')
            self.lineEdit.setText(self.lineEdit_PN.text())
        elif self.tabIndex == 1:
            self.tableWidget_DIDO.setVisible(False)
            self.tableWidget_AI.setVisible(True)
            self.tableWidget_AO.setVisible(False)
            self.tableWidget_CPU.setVisible(False)
            self.label_7.setText("AI")
            if self.comboBox_3.currentIndex()==0:
                self.lineEdit_PN.setText('P0000010392361')
            self.lineEdit_22.setText(self.lineEdit_PN.text())
        elif self.tabIndex == 2:
            self.tableWidget_DIDO.setVisible(False)
            self.tableWidget_AI.setVisible(False)
            self.tableWidget_AO.setVisible(True)
            self.tableWidget_CPU.setVisible(False)
            self.label_7.setText("AO")
            if self.comboBox_5.currentIndex()==0:
                self.lineEdit_PN.setText('P0000010392365')
            self.lineEdit_47.setText(self.lineEdit_PN.text())
        elif self.tabIndex == 3:
            self.tableWidget_DIDO.setVisible(False)
            self.tableWidget_AI.setVisible(False)
            self.tableWidget_AO.setVisible(False)
            self.tableWidget_CPU.setVisible(True)
            self.label_7.setText("CPU")
            if self.comboBox_20.currentIndex()==0:
                self.lineEdit_PN.setText('P0000010390631')
            self.lineEdit_30.setText(self.lineEdit_PN.text())
        self.saveDir = self.label_41.text()
        self.tabWidget.currentChanged.connect(self.tabChange)

        self.pushButton_9.setEnabled(True)
        self.pushButton_9.clicked.connect(self.uiResize)

        # self.checkBox_27.setChecked(False)
        self.checkBox_27.stateChanged.connect(self.DIDOCANAddr_stateChanged)
        self.lineEdit_6.setEnabled(self.checkBox_27.isChecked())
        self.label_16.setEnabled(self.checkBox_27.isChecked())
        self.lineEdit_7.setEnabled(self.checkBox_27.isChecked())
        self.label_17.setEnabled(self.checkBox_27.isChecked())

        # self.checkBox_28.setChecked(False)
        self.checkBox_28.stateChanged.connect(self.DIDOAdPara_stateChanged)
        self.lineEdit_13.setEnabled(self.checkBox_28.isChecked())
        self.label_22.setEnabled(self.checkBox_28.isChecked())
        self.lineEdit_12.setEnabled(self.checkBox_28.isChecked())
        self.label_23.setEnabled(self.checkBox_28.isChecked())
        self.lineEdit_14.setEnabled(self.checkBox_28.isChecked())
        self.label_26.setEnabled(self.checkBox_28.isChecked())
        self.label_25.setEnabled(self.checkBox_28.isChecked())
        self.label_24.setEnabled(self.checkBox_28.isChecked())

        # self.checkBox_3.setChecked(False)
        self.checkBox_3.stateChanged.connect(self.AIAdPara_stateChanged)
        self.lineEdit_15.setEnabled(self.checkBox_3.isChecked())
        self.label_42.setEnabled(self.checkBox_3.isChecked())
        self.lineEdit_16.setEnabled(self.checkBox_3.isChecked())
        self.label_43.setEnabled(self.checkBox_3.isChecked())
        self.lineEdit_17.setEnabled(self.checkBox_3.isChecked())
        self.label_44.setEnabled(self.checkBox_3.isChecked())
        self.label_45.setEnabled(self.checkBox_3.isChecked())
        self.label_46.setEnabled(self.checkBox_3.isChecked())

        # self.checkBox_4.setChecked(False)
        self.checkBox_4.stateChanged.connect(self.AICANAddr_stateChanged)
        self.lineEdit_18.setEnabled(self.checkBox_4.isChecked())
        self.label_48.setEnabled(self.checkBox_4.isChecked())
        self.lineEdit_19.setEnabled(self.checkBox_4.isChecked())
        self.label_49.setEnabled(self.checkBox_4.isChecked())
        self.lineEdit_23.setEnabled(self.checkBox_4.isChecked())
        self.label_54.setEnabled(self.checkBox_4.isChecked())

        # self.radioButton.setChecked(True)
        self.radioButton.toggled.connect(self.AI_NotCalibrate)

        # self.radioButton_2.setChecked(False)
        self.radioButton_2.toggled.connect(self.AI_Calibrate)

        # self.radioButton_3.setChecked(False)
        self.radioButton_3.toggled.connect(self.AI_Calibrate)

        # self.radioButton_4.setChecked(True)
        self.radioButton_4.toggled.connect(self.AI_NotTest)

        # self.radioButton_5.setChecked(False)
        self.radioButton_5.toggled.connect(self.AI_Test)

        # self.radioButton_16.setChecked(True)
        self.radioButton_16.toggled.connect(self.AO_NotTest)

        self.horizontalLayout_52.addWidget(self.radioButton_17)
        self.radioButton_17.toggled.connect(self.AO_Test)

        # self.checkBox_33.setChecked(False)
        self.checkBox_33.stateChanged.connect(self.AOCANAddr_stateChanged)
        self.lineEdit_39.setEnabled(self.checkBox_33.isChecked())
        self.label_74.setEnabled(self.checkBox_33.isChecked())
        self.lineEdit_40.setEnabled(self.checkBox_33.isChecked())
        self.label_75.setEnabled(self.checkBox_33.isChecked())
        self.lineEdit_41.setEnabled(self.checkBox_33.isChecked())
        self.label_76.setEnabled(self.checkBox_33.isChecked())

        # self.checkBox_34.setChecked(False)
        self.checkBox_34.stateChanged.connect(self.AOAdPara_stateChanged)
        self.lineEdit_42.setEnabled(self.checkBox_34.isChecked())
        self.label_77.setEnabled(self.checkBox_34.isChecked())
        self.lineEdit_43.setEnabled(self.checkBox_34.isChecked())
        self.label_78.setEnabled(self.checkBox_34.isChecked())
        self.lineEdit_44.setEnabled(self.checkBox_34.isChecked())
        self.label_79.setEnabled(self.checkBox_34.isChecked())
        self.label_80.setEnabled(self.checkBox_34.isChecked())
        self.label_81.setEnabled(self.checkBox_34.isChecked())

        self.radioButton_18.toggled.connect(self.AO_NotCalibrate)

        # self.radioButton_19.setChecked(True)
        self.radioButton_19.toggled.connect(self.AO_Calibrate)

        self.radioButton_20.toggled.connect(self.AO_Calibrate)

        # self.pushButton_9.clicked.connect(self.sendMessage)
        self.pushbutton_allScreen.setStyleSheet(f"""
                                    QPushButton {{
                                        image: url({self.current_dir}/全屏按钮.png);

                                    }}
                                    QPushButton:hover {{
                                        image: url({self.current_dir}/全屏按钮2.png);
                                    }}
                                """)
        self.pushbutton_cancelAllScreen.setStyleSheet(f"""
                                     QPushButton {{
                                         image: url({self.current_dir}/取消全屏按钮.png);

                                     }}
                                     QPushButton:hover {{
                                         image: url({self.current_dir}/取消全屏按钮2.png);
                                     }}
                                 """)
        self.pushButton_8.setStyleSheet(f"""
                                     QPushButton {{
                                         font: 75 26pt \"Arial\";
                                         background-color:  #11826C;
                                         color: rgb(255, 255, 255);
                                     }}
                                     QPushButton:hover {{
                                         font: 75 26pt \"Arial\";
                                         background-color:  #F5BD3D;
                                         color: rgb(0, 0, 0);

                                     }}
                                 """)
        self.pushButton_11.setStyleSheet(f"""
                                             QPushButton {{
                                                 font: 75 26pt \"Arial\";
                                                 background-color:  #11826C;
                                                 color: rgb(255, 255, 255);
                                             }}
                                             QPushButton:hover {{
                                                 font: 75 26pt \"Arial\";
                                                 background-color:  #F5BD3D;
                                                 color: rgb(0, 0, 0);

                                             }}
                                         """)


        self.pushbutton_allScreen.clicked.connect(self.allScreen)

        self.pushbutton_cancelAllScreen.setEnabled(False)
        self.pushbutton_cancelAllScreen.setVisible(False)
        self.pushbutton_cancelAllScreen.clicked.connect(self.cancelAllScreen)


        #CPU页面初始化
        self.CPU_comboBox_array = [self.comboBox_20,self.comboBox_21,self.comboBox_22,self.comboBox_23]
        for comboBox in self.CPU_comboBox_array:
            comboBox.currentIndexChanged.connect(self.saveConfig)
        for lineEdit in self.CPU_lineEdit_array:
            lineEdit.textChanged.connect(self.saveConfig)
        for checkBox in self.CPU_checkBox_array:
            checkBox.toggled.connect(self.saveConfig)
        self.CPU_paramChanged()
        self.checkBox_71.stateChanged.connect(self.CPU_paramChanged)
        self.checkBox_73.stateChanged.connect(self.testAllorNot)
        self.lineEdit_33.setText('00-E0-C0-4C-32-03-8B')


        # self.lineEdit_PN.setPlaceholderText('请输入PN码')
        self.lineEdit_SN.setPlaceholderText('请输入SN码')
        # self.lineEdit_SN.setReadOnly(True)
        self.lineEdit_REV.setPlaceholderText('请输入REV码')
        self.lineEdit_REV.setReadOnly(True)
        self.lineEdit_MAC.setPlaceholderText('请输入MAC地址')
        self.lineEdit_MAC.setReadOnly(True)
        # self.lineEdit_PN.setFocus()
        # self.lineEdit_PN.editingFinished.connect(self.inputPN)
        self.lineEdit_SN.setFocus()
        self.lineEdit_SN.editingFinished.connect(self.inputSN)
        self.lineEdit_REV.editingFinished.connect(self.inputREV)
        self.pushButton_10.clicked.connect(self.reInputPNSNREV)

        self.lineEdit_PN.textChanged.connect(self.changeTabWidgetPN)

        self.radioButton_16.toggled.connect(self.setFocusLineEdit)
        self.radioButton_17.toggled.connect(self.setFocusLineEdit)
        self.radioButton_18.toggled.connect(self.setFocusLineEdit)
        self.radioButton_19.toggled.connect(self.setFocusLineEdit)
        self.radioButton_20.toggled.connect(self.setFocusLineEdit)
        self.checkBox_29.toggled.connect(self.setFocusLineEdit)
        self.checkBox_30.toggled.connect(self.setFocusLineEdit)
        self.checkBox_31.toggled.connect(self.setFocusLineEdit)
        self.checkBox_32.toggled.connect(self.setFocusLineEdit)
        self.checkBox_36.toggled.connect(self.setFocusLineEdit)
        self.checkBox_35.toggled.connect(self.setFocusLineEdit)

        self.radioButton.toggled.connect(self.saveConfig)
        self.radioButton_2.toggled.connect(self.saveConfig)
        self.radioButton_3.toggled.connect(self.saveConfig)
        self.radioButton_4.toggled.connect(self.saveConfig)
        self.radioButton_5.toggled.connect(self.saveConfig)
        self.radioButton_16.toggled.connect(self.saveConfig)
        self.radioButton_17.toggled.connect(self.saveConfig)
        self.radioButton_18.toggled.connect(self.saveConfig)
        self.radioButton_19.toggled.connect(self.saveConfig)
        self.radioButton_20.toggled.connect(self.saveConfig)
        self.checkBox_29.toggled.connect(self.saveConfig)
        self.checkBox_30.toggled.connect(self.saveConfig)
        self.checkBox_31.toggled.connect(self.saveConfig)
        self.checkBox_32.toggled.connect(self.saveConfig)
        self.checkBox_36.toggled.connect(self.saveConfig)
        self.checkBox_35.toggled.connect(self.saveConfig)
        self.checkBox.toggled.connect(self.saveConfig)
        self.checkBox_2.toggled.connect(self.saveConfig)
        self.checkBox_5.toggled.connect(self.saveConfig)
        self.checkBox_6.toggled.connect(self.saveConfig)
        self.checkBox_7.toggled.connect(self.saveConfig)
        self.checkBox_8.toggled.connect(self.saveConfig)
        self.checkBox_9.toggled.connect(self.saveConfig)
        self.checkBox_10.toggled.connect(self.saveConfig)
        self.comboBox.currentIndexChanged.connect(self.saveConfig)
        self.comboBox_3.currentIndexChanged.connect(self.saveConfig)
        self.comboBox_5.currentIndexChanged.connect(self.saveConfig)
        self.comboBox.currentIndexChanged.connect(self.changePN)
        self.comboBox_3.currentIndexChanged.connect(self.changePN)
        self.comboBox_5.currentIndexChanged.connect(self.changePN)

        self.lineEdit_6.textChanged.connect(self.saveConfig)
        self.lineEdit_7.textChanged.connect(self.saveConfig)
        self.lineEdit_12.textChanged.connect(self.saveConfig)
        self.lineEdit_13.textChanged.connect(self.saveConfig)
        self.lineEdit_14.textChanged.connect(self.saveConfig)
        self.lineEdit_15.textChanged.connect(self.saveConfig)
        self.lineEdit_16.textChanged.connect(self.saveConfig)
        self.lineEdit_17.textChanged.connect(self.saveConfig)
        self.lineEdit_18.textChanged.connect(self.saveConfig)
        self.lineEdit_19.textChanged.connect(self.saveConfig)
        self.lineEdit_23.textChanged.connect(self.saveConfig)
        self.lineEdit_39.textChanged.connect(self.saveConfig)
        self.lineEdit_40.textChanged.connect(self.saveConfig)
        self.lineEdit_41.textChanged.connect(self.saveConfig)
        self.lineEdit_42.textChanged.connect(self.saveConfig)
        self.lineEdit_43.textChanged.connect(self.saveConfig)
        self.lineEdit_44.textChanged.connect(self.saveConfig)
        self.tabWidget.currentChanged.connect(self.saveConfig)
        self.pushButton.clicked.connect(self.saveConfig)
        self.pushButton_2.clicked.connect(self.saveConfig)
        self.pushButton_3.clicked.connect(self.saveConfig)
        self.pushButton_4.clicked.connect(self.saveConfig)

        self.pushButton_12.setEnabled(False)
        self.pushButton_12.setVisible(False)
        self.pushButton_12.clicked.connect(self.uiRecovery)
        self.pushButton_13.clicked.connect(self.option_pushButton13)

        #更新串口
        self.pushButton_renewSerial.clicked.connect(lambda:self.getSerialInf(1))
        global isMainRunning
        isMainRunning = True

    def generateDefaultConfigFile(self):
        config_str = "{'savePath': 'D:/MyData/wujun89/Desktop/EX300_x64_python'," \
                     "'currentIndex': 0," \
                     "'AO_型号': 0," \
                     "'AO_CAN_修改参数': False," \
                     "'AO_CAN_工装': '1'," \
                     "'AO_CAN_待检': '2'," \
                     "'AO_CAN_继电器': '3'," \
                     "'AO_附加_修改参数': False," \
                     "'AO_附加_波特率': '1000'," \
                     "'AO_附加_等待时间': '5000'," \
                     "'AO_附加_接收次数': '8'," \
                     "'AO_不标定': False," \
                     "'AO_仅标定1': False," \
                     "'AO_标定all': False," \
                     "'AO_标定电压': False," \
                     "'AO_标定电流': False," \
                     "'AO_不检测': False," \
                     "'AO_检测': False," \
                     "'AO_检测电压': False," \
                     "'AO_检测电流': False," \
                     "'AO_检测CAN': False," \
                     "'AO_检测run': False," \
                     "'AI_型号': 0," \
                     "'AI_CAN_修改参数': False," \
                     "'AI_CAN_工装': '1'," \
                     "'AI_CAN_待检': '2'," \
                     "'AI_CAN_继电器': '3'," \
                     "'AI_附加_修改参数': False," \
                     "'AI_附加_波特率': '1000'," \
                     "'AI_附加_等待时间': '5000'," \
                     "'AI_附加_接收次数': '8'," \
                     "'AI_不标定': False," \
                     "'AI_仅标定1': False," \
                     "'AI_标定all': False," \
                     "'AI_标定电压': False," \
                     "'AI_标定电流': False," \
                     "'AI_不检测': False," \
                     "'AI_检测': False," \
                     "'AI_检测电压': False," \
                     "'AI_检测电流': False," \
                     "'AI_检测CAN': False," \
                     "'AI_检测run': False," \
                     "'DIDO_型号': 0," \
                     "'DIDO_CAN_修改参数': False," \
                     "'DIDO_CAN_工装': '1'," \
                     "'DIDO_CAN_待检': '2'," \
                     "'DIDO_附加_修改参数': False," \
                     "'DIDO_附加_波特率': '1000'," \
                     "'DIDO_附加_间隔时间': '5000'," \
                     "'DIDO_附加_循环次数': '1'," \
                     "'DIDO_检测CAN': False," \
                     "'DIDO_检测run': False," \
                     "'CPU_型号': 0," \
                     "'CPU_IP': '192.168.1.55'," \
                     "'CPU_232COM': 0," \
                     "'CPU_485COM': 0," \
                     "'CPU_typecCOM': 0," \
                     "'型号检查': False," \
                     "'SRAM': False," \
                     "'FLASH': False," \
                     "'MAC/三码写入': False," \
                     "'FPGA': False," \
                     "'拨杆测试': False," \
                     "'MFK按键': False," \
                     "'掉电保存': False," \
                     "'RTC测试': False," \
                     "'各指示灯': False," \
                     "'本体IN': False," \
                     "'本体OUT': False," \
                     "'以太网': False," \
                     "'RS-232C': False," \
                     "'RS-485': False," \
                     "'右扩CAN': False," \
                     "'MA0202': False," \
                     "'测试报告': False," \
                     "'外观检测': False," \
                     "'工装1': '1'," \
                     "'工装2': '2'," \
                     "'工装3': '3'," \
                     "'工装4': '4'," \
                     "'工装5': '5'," \
                     "'修改参数': False}"
        self.configFile = open(f'{self.current_dir}/config.txt', 'w', encoding='utf-8')
        self.configFile.write(config_str)
        self.configFile.close()
    def loadConfig(self):
        try:
            with open(f'{self.current_dir}/config.txt','r+',encoding='utf-8') as file:
                config_content = file.read()
        except:
            self.showInf("配置文件不存在，初始化界面！"+self.HORIZONTAL_LINE)
            self.generateConfigFile()
            with open(f'{self.current_dir}/config.txt','r+',encoding='utf-8') as file:
                config_content = file.read()

        self.config_param = eval(config_content)
        self.label_41.setText(self.config_param["savePath"])
        self.tabWidget.setCurrentIndex(self.config_param["currentIndex"])

        # AO页面配置
        self.comboBox_5.setCurrentIndex(self.config_param["AO_型号"])
        self.checkBox_33.setChecked(self.config_param["AO_CAN_修改参数"])
        self.lineEdit_39.setText(self.config_param["AO_CAN_工装"])
        self.lineEdit_40.setText(self.config_param["AO_CAN_待检"])
        self.lineEdit_41.setText(self.config_param["AO_CAN_继电器"])
        self.checkBox_34.setChecked(self.config_param["AO_附加_修改参数"])
        self.lineEdit_42.setText(self.config_param["AO_附加_波特率"])
        self.lineEdit_43.setText(self.config_param["AO_附加_等待时间"])
        self.lineEdit_44.setText(self.config_param["AO_附加_接收次数"])
        self.radioButton_18.setChecked(self.config_param["AO_不标定"])
        self.radioButton_19.setChecked(self.config_param["AO_仅标定1"])
        self.radioButton_20.setChecked(self.config_param["AO_标定all"])
        self.checkBox_35.setChecked(self.config_param["AO_标定电压"])
        self.checkBox_36.setChecked(self.config_param["AO_标定电流"])
        self.radioButton_16.setChecked(self.config_param["AO_不检测"])
        self.radioButton_17.setChecked(self.config_param["AO_检测"])
        self.checkBox_29.setChecked(self.config_param["AO_检测电压"])
        self.checkBox_30.setChecked(self.config_param["AO_检测电流"])
        self.checkBox_31.setChecked(self.config_param["AO_检测CAN"])
        self.checkBox_32.setChecked(self.config_param["AO_检测run"])
        # AI页面配置
        self.comboBox_3.setCurrentIndex(self.config_param["AI_型号"])
        self.checkBox_4.setChecked(self.config_param["AI_CAN_修改参数"])
        self.lineEdit_18.setText(self.config_param["AI_CAN_工装"])
        self.lineEdit_19.setText(self.config_param["AI_CAN_待检"])
        self.lineEdit_23.setText(self.config_param["AI_CAN_继电器"])
        self.checkBox_3.setChecked(self.config_param["AI_附加_修改参数"])
        self.lineEdit_16.setText(self.config_param["AI_附加_波特率"])
        self.lineEdit_17.setText(self.config_param["AI_附加_等待时间"])
        self.lineEdit_15.setText(self.config_param["AI_附加_接收次数"])
        self.radioButton.setChecked(self.config_param["AI_不标定"])
        self.radioButton_2.setChecked(self.config_param["AI_仅标定1"])
        self.radioButton_3.setChecked(self.config_param["AI_标定all"])
        self.checkBox_5.setChecked(self.config_param["AI_标定电压"])
        self.checkBox_6.setChecked(self.config_param["AI_标定电流"])
        self.radioButton_4.setChecked(self.config_param["AI_不检测"])
        self.radioButton_5.setChecked(self.config_param["AI_检测"])
        self.checkBox_9.setChecked(self.config_param["AI_检测电压"])
        self.checkBox_10.setChecked(self.config_param["AI_检测电流"])
        self.checkBox_8.setChecked(self.config_param["AI_检测CAN"])
        self.checkBox_7.setChecked(self.config_param["AI_检测run"])
        # DIDO页面配置
        self.comboBox.setCurrentIndex(self.config_param["DIDO_型号"])
        self.checkBox_27.setChecked(self.config_param["DIDO_CAN_修改参数"])
        self.lineEdit_6.setText(self.config_param["DIDO_CAN_工装"])
        self.lineEdit_7.setText(self.config_param["DIDO_CAN_待检"])
        self.checkBox_28.setChecked(self.config_param["DIDO_附加_修改参数"])
        self.lineEdit_13.setText(self.config_param["DIDO_附加_波特率"])
        self.lineEdit_12.setText(self.config_param["DIDO_附加_间隔时间"])
        self.lineEdit_14.setText(self.config_param["DIDO_附加_循环次数"])
        self.checkBox.setChecked(self.config_param["DIDO_检测CAN"])
        self.checkBox_2.setChecked(self.config_param["DIDO_检测run"])

        # CPU页面配置
        self.comboBox_20.setCurrentIndex(self.config_param["CPU_型号"])
        self.comboBox_21.setCurrentIndex(self.config_param["CPU_232COM"])
        self.comboBox_22.setCurrentIndex(self.config_param["CPU_485COM"])
        self.comboBox_23.setCurrentIndex(self.config_param["CPU_typecCOM"])

        for i in range(len(self.CPU_lineEdit_array)):
            self.CPU_lineEdit_array[i].setText(self.config_param[self.CPU_lineEditName_array[i]])


        for i in range(len(self.CPU_checkBox_array)):
            self.CPU_checkBox_array[i].setChecked(self.config_param[self.CPU_checkBoxName_array[i]])

    def saveConfig(self):
        self.config_param["savePath"] = self.label_41.text()
        self.config_param["currentIndex"] = self.tabWidget.currentIndex()

        # AO页面配置
        self.config_param["AO_型号"] = self.comboBox_5.currentIndex()
        self.config_param["AO_CAN_修改参数"] = self.checkBox_33.isChecked()
        self.config_param["AO_CAN_工装"] = self.lineEdit_39.text()
        self.config_param["AO_CAN_待检"] = self.lineEdit_40.text()
        self.config_param["AO_CAN_继电器"] = self.lineEdit_41.text()
        self.config_param["AO_附加_修改参数"] = self.checkBox_34.isChecked()
        self.config_param["AO_附加_波特率"] = self.lineEdit_42.text()
        self.config_param["AO_附加_等待时间"] = self.lineEdit_43.text()
        self.config_param["AO_附加_接收次数"] = self.lineEdit_44.text()
        self.config_param["AO_不标定"] = self.radioButton_18.isChecked()
        self.config_param["AO_仅标定1"] = self.radioButton_19.isChecked()
        self.config_param["AO_标定all"] = self.radioButton_20.isChecked()
        self.config_param["AO_标定电压"] = self.checkBox_35.isChecked()
        self.config_param["AO_标定电流"] = self.checkBox_36.isChecked()
        self.config_param["AO_不检测"] = self.radioButton_16.isChecked()
        self.config_param["AO_检测"] = self.radioButton_17.isChecked()
        self.config_param["AO_检测电压"] = self.checkBox_29.isChecked()
        self.config_param["AO_检测电流"] = self.checkBox_30.isChecked()
        self.config_param["AO_检测CAN"] = self.checkBox_31.isChecked()
        self.config_param["AO_检测run"] = self.checkBox_32.isChecked()
        # AI页面配置
        self.config_param["AI_型号"] = self.comboBox_3.currentIndex()
        self.config_param["AI_CAN_修改参数"] = self.checkBox_4.isChecked()
        self.config_param["AI_CAN_工装"] = self.lineEdit_18.text()
        self.config_param["AI_CAN_待检"] = self.lineEdit_19.text()
        self.config_param["AI_CAN_继电器"] = self.lineEdit_23.text()
        self.config_param["AI_附加_修改参数"] = self.checkBox_3.isChecked()
        self.config_param["AI_附加_波特率"] = self.lineEdit_16.text()
        self.config_param["AI_附加_等待时间"] = self.lineEdit_17.text()
        self.config_param["AI_附加_接收次数"] = self.lineEdit_15.text()
        self.config_param["AI_不标定"] = self.radioButton.isChecked()
        self.config_param["AI_仅标定1"] = self.radioButton_2.isChecked()
        self.config_param["AI_标定all"] = self.radioButton_3.isChecked()
        self.config_param["AI_标定电压"] = self.checkBox_5.isChecked()
        self.config_param["AI_标定电流"] = self.checkBox_6.isChecked()
        self.config_param["AI_不检测"] = self.radioButton_4.isChecked()
        self.config_param["AI_检测"] = self.radioButton_5.isChecked()
        self.config_param["AI_检测电压"] = self.checkBox_9.isChecked()
        self.config_param["AI_检测电流"] = self.checkBox_10.isChecked()
        self.config_param["AI_检测CAN"] = self.checkBox_8.isChecked()
        self.config_param["AI_检测run"] = self.checkBox_7.isChecked()
        # DIDO页面配置
        self.config_param["DIDO_型号"] = self.comboBox.currentIndex()
        self.config_param["DIDO_CAN_修改参数"] = self.checkBox_27.isChecked()
        self.config_param["DIDO_CAN_工装"] = self.lineEdit_6.text()
        self.config_param["DIDO_CAN_待检"] = self.lineEdit_7.text()
        self.config_param["DIDO_附加_修改参数"] = self.checkBox_28.isChecked()
        self.config_param["DIDO_附加_波特率"] = self.lineEdit_13.text()
        self.config_param["DIDO_附加_间隔时间"] = self.lineEdit_12.text()
        self.config_param["DIDO_附加_循环次数"] = self.lineEdit_14.text()
        self.config_param["DIDO_检测CAN"] = self.checkBox.isChecked()
        self.config_param["DIDO_检测run"] = self.checkBox_2.isChecked()

        #CPU页面配置
        for i in range(len(self.CPU_lineEdit_array)):
            self.config_param[self.CPU_lineEditName_array[i]] = self.CPU_lineEdit_array[i].text()
        self.config_param["CPU_型号"] = self.comboBox_20.currentIndex()
        self.config_param["CPU_232COM"] = self.comboBox_21.currentIndex()
        self.config_param["CPU_485COM"] = self.comboBox_22.currentIndex()
        self.config_param["CPU_typecCOM"] = self.comboBox_23.currentIndex()
        for i in range(len(self.CPU_checkBox_array)):
            self.config_param[self.CPU_checkBoxName_array[i]] = self.CPU_checkBox_array[i].isChecked()

        #save
        config_str = str(self.config_param)
        config_str1=''
        pos=0
        while pos>=0:
            pos=config_str.find(',')
            config_str1=config_str1+config_str[:pos+1]+'\n'
            config_str=config_str[pos+1:]
        config_str=config_str1+config_str
        self.configFile = open(f'{self.current_dir}/config.txt', 'w',encoding='utf-8')
        self.configFile.write(config_str)
        self.configFile.close()

    def changePN(self):
        if self.tabIndex == 0:
            if self.comboBox.currentIndex() == 1:
                self.lineEdit_PN.setText('P0000010391877')
            elif self.comboBox.currentIndex() == 4:
                self.lineEdit_PN.setText('P0000010392086')
            elif self.comboBox.currentIndex() == 5:
                self.lineEdit_PN.setText('P0000010392121')
        elif self.tabIndex == 1:
            if self.comboBox_3.currentIndex() == 0:
                self.lineEdit_PN.setText('P0000010392361')
        elif self.tabIndex == 2:
            if self.comboBox_5.currentIndex() == 0:
                self.lineEdit_PN.setText('P0000010392365')

    def changeTabWidgetPN(self):
        if self.tabIndex == 0:
            self.lineEdit.setText(self.lineEdit_PN.text())
        elif self.tabIndex == 1:
            self.lineEdit_22.setText(self.lineEdit_PN.text())
        elif self.tabIndex == 2:
            self.lineEdit_47.setText(self.lineEdit_PN.text())

    def update_time(self):
        self.label_47.setText(datetime.datetime.now().strftime('%m/%d %H:%M:%S'))
        QTimer.singleShot(1000, self.update_time)

    def killUi(self):
        self.saveConfig()
        self.close()
        try:
            os.kill(os.getpid(), signal.SIGTERM)
        except OSError as e:
            self.showInf(f"无法杀死当前进程: {e}\n")


    def label_mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.dragging = True
            self.offset = event.pos()

    def label_mouseMoveEvent(self, event):
        if self.dragging and event.buttons() & Qt.LeftButton:
            self.move(event.globalPos() - self.offset)
            event.accept()

    def mouseReleaseEvent(self, event):
        self.dragging = False

    def uiResize(self):
        self.resize(740,872)
        self.label_6.resize(740,30)
        self.pushButton_11.setGeometry(QtCore.QRect(680, 0, 30, 30))
        self.pushButton_8.setGeometry(QtCore.QRect(710, 0, 30, 30))

        self.pushButton_9.setEnabled(False)
        self.pushButton_9.setVisible(False)
        self.pushButton_12.setEnabled(True)
        self.pushButton_12.setVisible(True)

    def uiRecovery(self):
        self.resize(1440, 872)
        self.label_6.resize(1440, 30)
        self.pushButton_11.setGeometry(QtCore.QRect(1380, 0, 30, 30))
        self.pushButton_8.setGeometry(QtCore.QRect(1410, 0, 30, 30))

        self.pushButton_12.setEnabled(False)
        self.pushButton_12.setVisible(False)
        self.pushButton_9.setEnabled(True)
        self.pushButton_9.setVisible(True)

    def setFocusLineEdit(self):
        # if len(self.lineEdit_PN.text()) == 0 and len(self.lineEdit_SN.text()) == 0 and len(self.lineEdit_REV.text()) == 0:
        #     self.lineEdit_PN.setReadOnly(False)
        #     self.lineEdit_PN.setFocus()
        if len(self.lineEdit_SN.text()) == 0 and len(self.lineEdit_REV.text()) == 0:
            self.lineEdit_SN.setReadOnly(False)
            self.lineEdit_SN.setFocus()

    def reInputPNSNREV(self):
        # self.textBrowser_5.clear()
        # self.lineEdit_PN.clear()
        self.lineEdit_SN.clear()
        self.lineEdit_REV.clear()
        # self.lineEdit_PN.setReadOnly(False)
        self.lineEdit_SN.setReadOnly(False)
        self.lineEdit_REV.setReadOnly(True)
        self.lineEdit_SN.setFocus()

        # self.lineEdit.clear()
        self.lineEdit_3.clear()
        self.lineEdit_5.clear()

        # self.lineEdit_22.clear()
        self.lineEdit_21.clear()
        self.lineEdit_20.clear()

        self.lineEdit_31.clear()
        self.lineEdit_32.clear()

        # self.lineEdit_47.clear()
        self.lineEdit_46.clear()
        self.lineEdit_45.clear()
        self.pushButton_4.setEnabled(False)
        self.pushButton_4.setStyleSheet(self.topButton_qss['off'])


    def inputPN(self):
        if len(self.lineEdit_PN.text()) == 15:
            self.lineEdit_PN.setReadOnly(True)
            self.lineEdit_SN.setReadOnly(False)
            self.lineEdit_SN.setFocus()
            if self.tabIndex == 0:
                self.lineEdit.setText(self.lineEdit_PN.text())
            elif self.tabIndex == 1:
                self.lineEdit_22.setText(self.lineEdit_PN.text())
            elif self.tabIndex == 2:
                self.lineEdit_47.setText(self.lineEdit_PN.text())

        else:
            # time.sleep(0.1)
            self.lineEdit_PN.clear()

    def inputSN(self):
        # # print(self.lineEdit_PN.text())
        if len(self.lineEdit_SN.text()) == 12:
            self.lineEdit_SN.setReadOnly(True)
            self.lineEdit_REV.setReadOnly(False)
            self.lineEdit_REV.setFocus()
            if self.tabIndex == 0:
                self.lineEdit_3.setText(self.lineEdit_SN.text())
            elif self.tabIndex == 1:
                self.lineEdit_21.setText(self.lineEdit_SN.text())
            elif self.tabIndex == 2:
                self.lineEdit_46.setText(self.lineEdit_SN.text())
            elif self.tabIndex == 3:
                self.lineEdit_31.setText(self.lineEdit_SN.text())
        else:
            # time.sleep(0.1)
            self.lineEdit_SN.clear()

    def inputREV(self):
        if len(self.lineEdit_REV.text()) == 2:
            # self.showInf('三码完成输入')
            self.lineEdit_REV.setReadOnly(True)
            self.pushButton_4.setEnabled(True)
            self.pushButton_4.setStyleSheet(self.topButton_qss['on'])
            self.pushButton_4.setStyleSheet(f"""
                                                QPushButton {{
                                                    background-color: #11826C;
                                                    color:white;
                                                    font: bold 25pt \"宋体\"
                                                }}
                                                QPushButton:hover {{
                                                    background-color: {self.hover_color.name()};
                                                    color:#212B36;
                                                    font: bold 25pt \"宋体\"
                                                }}
                                            """)
            if self.tabIndex == 0:
                self.lineEdit_5.setText(self.lineEdit_REV.text())
            elif self.tabIndex == 1:
                self.lineEdit_20.setText(self.lineEdit_REV.text())
            elif self.tabIndex == 2:
                self.lineEdit_45.setText(self.lineEdit_REV.text())
            elif self.tabIndex == 3:
                self.lineEdit_32.setText(self.lineEdit_REV.text())
            self.pushButton_4.setFocus()
        else:
            # time.sleep(0.5)
            self.lineEdit_REV.clear()

    def changeSaveDir(self):
        self.saveDir = self.label_41.text()
        self.saveDir = QFileDialog.getExistingDirectory(self, '修改路径', self.saveDir)
        if self.saveDir !='':
            self.label_41.setText(self.saveDir)
            # print(self.saveDir)

    def openSaveDir(self):
        os.startfile(self.label_41.text())

    def tabChange(self):
        self.tabIndex = self.tabWidget.currentIndex()
        # print("current tabIndex:",self.tabIndex)
        if self.tabIndex == 0 and not self.isAllScreen:
            self.tableWidget_DIDO.setVisible(True)
            self.tableWidget_AI.setVisible(False)
            self.tableWidget_AO.setVisible(False)
            self.tableWidget_CPU.setVisible(False)
            self.label_7.setText("DIDO")
            if self.comboBox.currentIndex() == 1:
                self.lineEdit_PN.setText('P0000010391877')
            elif self.comboBox.currentIndex() == 4:
                self.lineEdit_PN.setText('P0000010392086')
            elif self.comboBox.currentIndex() == 5:
                self.lineEdit_PN.setText('P0000010392121')
        elif self.tabIndex == 1 and not self.isAllScreen:
            self.tableWidget_DIDO.setVisible(False)
            self.tableWidget_AI.setVisible(True)
            self.tableWidget_AO.setVisible(False)
            self.tableWidget_CPU.setVisible(False)
            self.label_7.setText("AI")
            if self.comboBox_3.currentIndex() == 0:
                self.lineEdit_PN.setText('P0000010392361')
        elif self.tabIndex == 2 and not self.isAllScreen:
            self.tableWidget_DIDO.setVisible(False)
            self.tableWidget_AI.setVisible(False)
            self.tableWidget_AO.setVisible(True)
            self.tableWidget_CPU.setVisible(False)
            self.label_7.setText("AO")
            if self.comboBox_5.currentIndex() == 0:
                self.lineEdit_PN.setText('P0000010392365')
        elif self.tabIndex == 3 and not self.isAllScreen:
            self.tableWidget_DIDO.setVisible(False)
            self.tableWidget_AI.setVisible(False)
            self.tableWidget_AO.setVisible(False)
            self.tableWidget_CPU.setVisible(True)
            self.label_7.setText("CPU")
            if self.comboBox_20.currentIndex() == 0:
                self.lineEdit_PN.setText('P0000010390631')
        # self.reInputPNSNREV()

    def start_button(self):

        self.label.setStyleSheet(self.testState_qss['testing'])
        self.label.setText('检测中……')
        self.pushButton_3.setEnabled(True)
        self.pushButton_3.setStyleSheet(self.topButton_qss['stop_on'])
        self.pushButton_3.setStyleSheet(f"""
                                            QPushButton {{
                                                {self.topButton_qss['stop_on']}
                                            }}
                                            QPushButton:hover {{
                                                background-color: rgb(255,0,0);
                                                color:rgb(0, 0, 0)
                                            }}
                                        """)
        self.pushButton_4.setEnabled(False)
        self.pushButton_4.setVisible(False)

        self.pushButton_pause.setStyleSheet(f"""
                                                QPushButton {{
                                                    {self.topButton_qss['pause_on']}
                                                }}
                                                QPushButton:hover {{
                                                    background-color: rgb(255, 255, 0);
                                                    color:rgb(0, 0, 0)
                                                }}
                                            """)
        self.pushButton_pause.setVisible(True)
        self.pushButton_pause.setEnabled(True)
        self.textBrowser_5.clear()
        global isMainRunning
        isMainRunning = True




        if not self.startTest():
            self.PassOrFail(False)
            self.showInf('检测已停止！\n\n')



    def pause_button(self):
        CAN_option.receivePause()
        # self.showInf(self.HORIZONTAL_LINE + '暂停中…………' + self.HORIZONTAL_LINE)
        self.label.setText('暂停中……')
        self.pushButton_pause.setVisible(False)
        self.pushButton_pause.setEnabled(False)
        self.pushButton_resume.setStyleSheet(self.topButton_qss['start_on'])
        self.pushButton_resume.setStyleSheet(f"""
                                QPushButton {{
                                    {self.topButton_qss['start_on']}
                                }}
                                QPushButton:hover {{
                                    background-color: rgb(0, 255, 0);
                                    color:rgb(0, 0, 0)
                                }}
                            """)
        self.pushButton_resume.setEnabled(True)
        self.pushButton_resume.setVisible(True)

        # self.work_thread.pause()


    def resume_button(self):
        CAN_option.receiveResume()
        self.label.setText('测试中……')
        self.pushButton_resume.setVisible(False)
        self.pushButton_resume.setEnabled(False)
        self.pushButton_pause.setStyleSheet(f"""
                                        QPushButton {{
                                            {self.topButton_qss['pause_on']}
                                        }}
                                        QPushButton:hover {{
                                            background-color: rgb(255, 255, 0);
                                            color:rgb(0, 0, 0)
                                        }}
                                    """)
        self.pushButton_pause.setEnabled(True)
        self.pushButton_pause.setVisible(True)
        # self.work_thread.bresume()

    def stop_button(self):
        # self.textBrowser_5.insertPlainText(self.HORIZONTAL_LINE + '停止测试\n' + self.HORIZONTAL_LINE)
        # self.move_to_end()
        self.showInf(self.HORIZONTAL_LINE + '停止测试\n' + self.HORIZONTAL_LINE)
        self.stop_signal.emit(True)
        self.pushButton_3.setEnabled(False)
        self.pushButton_3.setStyleSheet(self.topButton_qss['off'])
        self.pushButton_resume.setVisible(False)
        self.pushButton_resume.setEnabled(False)
        self.pushButton_pause.setEnabled(False)
        self.pushButton_pause.setVisible(False)
        self.pushButton_4.setEnabled(False)
        self.pushButton_4.setVisible(True)
        self.pushButton_4.setStyleSheet(self.topButton_qss['off'])
        self.reInputPNSNREV()
        # self.pushButton_9.setStyleSheet(self.topButton_qss['off'])
        # self.subRunFlag = False

        CAN_option.receiveStop()

        global isMainRunning
        isMainRunning = False

    def startTest(self):
        try:
            # 启动CAN分析仪
            # can_thread=threading.Thread(target = canInit_thread)
            # can_thread.start()
            # event_canInit.wait()
            # list_canInit = q.get()
            # if not list_canInit[0]:
            #     self.showMessageBox(list_canInit[1])
            #     self.CANFail()
            #     return False
            list_canInit=CAN_init(1)
            if not list_canInit[0]:
                self.showMessageBox(list_canInit[1])
                self.CANFail()
                return False

            CAN_option.receiveRun()
            CAN_option.receiveResume()
            self.pushButton_10.setEnabled(False)
            self.pushButton_10.setStyleSheet('color: rgb(255, 255, 255);background-color: rgb(197, 197, 197);')
            try:
                if not self.sendMessage():
                    return False
            except Exception as e:
                self.showInf(f"sendMessageError:{e}{self.HORIZONTAL_LINE}")
                # 捕获异常并输出详细的错误信息
                traceback.print_exc()
                return False

            QApplication.processEvents()

            #节点分配
            if self.tabIndex == 0:
                if not self.configCANAddr(int(self.lineEdit_6.text()), int(self.lineEdit_7.text()), '', '', ''):
                    return False
            elif self.tabIndex == 1:
                if not self.configCANAddr(int(self.lineEdit_18.text()),int(self.lineEdit_23.text()),
                                          int(self.lineEdit_23.text())+1, int(self.lineEdit_19.text()), ''):
                    return False
            elif self.tabIndex == 2:
                if not self.configCANAddr(int(self.lineEdit_39.text()),int(self.lineEdit_41.text()),
                                          int(self.lineEdit_41.text())+1, int(self.lineEdit_40.text()), ''):
                    return False
            elif self.tabIndex == 3:
                if not self.configCANAddr(int(self.lineEdit_34.text()), int(self.lineEdit_35.text()),
                                          int(self.lineEdit_36.text()), int(self.lineEdit_37.text()),
                                          int(self.lineEdit_38.text())):
                    return False

            self.result_queue = Queue()
            self.testFlag = ''

            # #读设备三码
            # bool_3code,code_list = get3codeFromPLC(2)
            # if bool_3code:
            #     QMessageBox.about(None, '通过', f'三码一致！\nPN:{code_list[0]}\nSN:{code_list[1]}\nREV:{code_list[2]}')
            # else:
            #     QMessageBox.critical(None, '警告', f'三码不一致！\nPN:{code_list[0]}\nSN:{code_list[1]}\nREV:{code_list[2]}',
            #                          QMessageBox.Yes | QMessageBox.No,
            #                          QMessageBox.Yes)
            #     return False



            #开始时间
            self.allStart_time = time.time()
            if not mainThreadRunning():
                return False

            if not mainThreadRunning():
                return False
            if len(self.module_pn) == 0 or len(self.module_sn) == 0 or len(self.module_rev) == 0:
                # self.isPause()
                if not mainThreadRunning():
                    return False
                reply = QMessageBox.warning(None, '警告', '产品三码信息不全，请重新扫入！',
                                            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                return False

            if self.tabWidget.currentIndex() == 0:
                self.testFlag = 'DIDO'
                self.appearanceTest(self.testFlag)
                mTable = self.tableWidget_DIDO
                if self.comboBox.currentIndex() == 0 or self.comboBox.currentIndex() == 1 \
                        or self.comboBox.currentIndex() == 2:
                    self.testFlag = 'DI'

                elif self.comboBox.currentIndex() == 3 or self.comboBox.currentIndex() == 4 \
                        or self.comboBox.currentIndex() == 5 or self.comboBox.currentIndex() == 6:
                    self.testFlag = 'DO'

                self.DIDO_thread = QThread()
                from thread_DIDO import DIDOThread
                self.DIDO_option = DIDOThread(self.inf_DIDOlist, self.result_queue, self.appearance, self.testFlag)
                self.DIDO_option.result_signal.connect(self.showInf)
                self.DIDO_option.item_signal.connect(self.DIDO_itemOperation)
                self.DIDO_option.pass_signal.connect(self.PassOrFail)
                # self.DIDO_option.RunErr_signal.connect(self.testRunErr)
                # self.DIDO_option.CANRunErr_signal.connect(self.testCANRunErr)
                self.DIDO_option.messageBox_signal.connect(self.showMessageBox)
                # self.DIDO_option.excel_signal.connect(self.generateExcel)
                self.DIDO_option.allFinished_signal.connect(self.allFinished)
                self.DIDO_option.label_signal.connect(self.labelChange)
                self.DIDO_option.saveExcel_signal.connect(self.saveExcel)
                # self.DIDO_option.print_signal.connect(self.printResult)

                self.pushButton_3.clicked.connect(self.DIDO_option.stop_work)
                self.pushButton_pause.clicked.connect(self.DIDO_option.pause_work)
                self.pushButton_resume.clicked.connect(self.DIDO_option.resume_work)

                self.DIDO_option.moveToThread(self.DIDO_thread)
                self.DIDO_thread.started.connect(self.DIDO_option.DIDOOption)
                self.DIDO_thread.start()




            if self.tabWidget.currentIndex() == 1:
                # self.lineEdit_PN = 'PRDr9HPA06Mz-00'
                # self.lineEdit_SN = 'S1223-001083'
                # self.lineEdit_REV = '06'

                self.testFlag = 'AI'
                self.appearanceTest(self.testFlag)
                # self.appearance = True
                # self.showInf(f'self.tabWidget.currentIndex()={self.tabWidget.currentIndex()}\n\n')
                mTable = self.tableWidget_AI
                # isPassVol = True
                # isPassCur = True
                if self.isTest:
                    if self.isTestRunErr:
                        self.testNum = self.testNum - 1
                        # self.testRunErr(self.CANAddr_AI)
                    elif not self.isTestRunErr:
                        if not mainThreadRunning():
                            return False
                        self.AI_itemOperation([1, 0, 0, ''])
                        if not mainThreadRunning():
                            return False
                        self.AI_itemOperation([2, 0, 0, ''])
                    # self.showInf(f'RunErrself.testNum = {self.testNum}\n\n')
                    if self.isTestCANRunErr:
                        self.testNum = self.testNum - 1
                        # self.testCANRunErr(self.CANAddr_AI)
                    elif not self.isTestCANRunErr:
                        if not mainThreadRunning():
                            return False
                        self.AI_itemOperation([3, 0, 0, ''])
                        if not mainThreadRunning():
                            return False
                        self.AI_itemOperation([4, 0, 0, ''])
                        # self.itemOperation(mTable, 3, 0, 0, '')
                        # self.itemOperation(mTable, 4, 0, 0, '')
                    if self.isAITestVol:
                        self.testNum = self.testNum - 1
                    elif not self.isAITestVol:
                        if not mainThreadRunning():
                            return False
                        self.AI_itemOperation([7, 0, 0, ''])
                    if self.isAITestCur:
                        self.testNum = self.testNum - 1
                    elif not self.isAITestCur:
                        if not mainThreadRunning():
                            return False
                        self.AI_itemOperation([8, 0, 0, ''])



                # if self.isCalibrate:
                    # self.calibrateAI(mTable)
                self.AI_thread = None
                self.worker = None
                if not self.AI_thread or not self.worker_thread.isRunning():
                    # # 创建队列用于主线程和子线程之间的通信
                    # self.result_queue = Queue()

                    self.AI_thread = QThread()
                    from thread_AI import AIThread
                    self.AI_option = AIThread(self.inf_AIlist,self.result_queue,self.appearance)

                    self.AI_option.result_signal.connect(self.showInf)
                    self.AI_option.item_signal.connect(self.AI_itemOperation)
                    self.AI_option.pass_signal.connect(self.PassOrFail)
                    # self.AI_option.RunErr_signal.connect(self.testRunErr)
                    # self.AI_option.CANRunErr_signal.connect(self.testCANRunErr)
                    self.AI_option.messageBox_signal.connect(self.showMessageBox)
                    # self.AI_option.excel_signal.connect(self.generateExcel)
                    self.AI_option.allFinished_signal.connect(self.allFinished)
                    self.AI_option.label_signal.connect(self.labelChange)
                    self.AI_option.saveExcel_signal.connect(self.saveExcel)#保存测试报告
                    # self.AI_option.print_signal.connect(self.printResult)#打印测试标签

                    self.pushButton_3.clicked.connect(self.AI_option.stop_work)
                    self.pushButton_pause.clicked.connect(self.AI_option.pause_work)
                    self.pushButton_resume.clicked.connect(self.AI_option.resume_work)

                    self.AI_option.moveToThread(self.AI_thread)
                    self.AI_thread.started.connect(self.AI_option.AIOption)
                    self.AI_thread.start()



            elif self.tabWidget.currentIndex() == 2:
                self.testFlag = 'AO'
                mTable = self.tableWidget_AO

                self.appearanceTest(self.testFlag)
                self.AO_thread = None
                self.worker = None
                if not self.AO_thread or not self.worker_thread.isRunning():
                    # # 创建队列用于主线程和子线程之间的通信
                    # self.result_queue = Queue()

                    self.AO_thread = QThread()
                    from thread_AO import AOThread
                    self.AO_option = AOThread(self.inf_AOlist, self.result_queue, self.appearance)
                    self.AO_option.result_signal.connect(self.showInf)
                    self.AO_option.item_signal.connect(self.AO_itemOperation)
                    self.AO_option.pass_signal.connect(self.PassOrFail)
                    # self.AO_option.RunErr_signal.connect(self.testRunErr)
                    # self.AO_option.CANRunErr_signal.connect(self.testCANRunErr)
                    self.AO_option.messageBox_signal.connect(self.showMessageBox)
                    # self.AO_option.excel_signal.connect(self.generateExcel)
                    self.AO_option.allFinished_signal.connect(self.allFinished)
                    self.AO_option.label_signal.connect(self.labelChange)
                    self.AO_option.saveExcel_signal.connect(self.saveExcel)
                    # self.AO_option.print_signal.connect(self.printResult)

                    self.pushButton_3.clicked.connect(self.AO_option.stop_work)
                    self.pushButton_pause.clicked.connect(self.AO_option.pause_work)
                    self.pushButton_resume.clicked.connect(self.AO_option.resume_work)

                    self.AO_option.moveToThread(self.AO_thread)
                    self.AO_thread.started.connect(self.AO_option.AOOption)
                    self.AO_thread.start()

            elif self.tabWidget.currentIndex() == 3:
                self.testFlag = 'CPU'
                mTable = self.tableWidget_CPU
                #CPU的外观检测选项放在子线程里进行并且可选
                #self.appearanceTest(self.testFlag)
                self.CPU_thread = None
                self.worker = None
                if not self.CPU_thread or not self.worker_thread.isRunning():
                    # # 创建队列用于主线程和子线程之间的通信
                    # self.result_queue = Queue()

                    self.CPU_thread = QThread()
                    from thread_CPU import CPUThread
                    self.CPU_option = CPUThread(self.inf_CPUlist, self.result_queue)
                    self.CPU_option.result_signal.connect(self.showInf)
                    self.CPU_option.item_signal.connect(self.CPU_itemOperation)
                    self.CPU_option.pass_signal.connect(self.PassOrFail)
                    # self.CPU_option.RunErr_signal.connect(self.testRunErr)
                    # self.CPU_option.CANRunErr_signal.connect(self.testCANRunErr)
                    self.CPU_option.messageBox_signal.connect(self.showMessageBox)
                    # self.CPU_option.excel_signal.connect(self.generateExcel)
                    self.CPU_option.allFinished_signal.connect(self.allFinished)
                    self.CPU_option.label_signal.connect(self.labelChange)
                    self.CPU_option.saveExcel_signal.connect(self.saveExcel)
                    # self.CPU_option.print_signal.connect(self.printResult)

                    self.pushButton_3.clicked.connect(self.CPU_option.stop_work)
                    self.pushButton_pause.clicked.connect(self.CPU_option.pause_work)
                    self.pushButton_resume.clicked.connect(self.CPU_option.resume_work)

                    self.CPU_option.moveToThread(self.CPU_thread)
                    self.CPU_thread.started.connect(self.CPU_option.CPUOption)
                    self.CPU_thread.start()


        except Exception as e:
            self.showInf(f"startTestError:{e}{self.HORIZONTAL_LINE}")
            # 捕获异常并输出详细的错误信息
            traceback.print_exc()
            return False

        return True


    def appearanceTest(self,testFlag):
        if testFlag == 'DIDO':
            mTable = self.tableWidget_DIDO
        elif testFlag == 'AI':
            mTable = self.tableWidget_AI
        elif testFlag == 'AO':
            mTable = self.tableWidget_AO
        # elif testFlag == 'CPU':
        #     mTable = self.tableWidget_CPU
        # 检测外观
        appearanceStart_time = time.time()
        if not mainThreadRunning():
            return False

        self.itemOperation(mTable,0, 1, 0, '')

        if not mainThreadRunning():
            return False
        reply = QMessageBox.question(None, '外观检测', '产品外观是否完好?',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            self.appearance = True
            appearanceEnd_time = time.time()
            appearanceTest_time = round(appearanceEnd_time - appearanceStart_time, 2)
            if not mainThreadRunning():
                return False
            self.itemOperation(mTable, 0, 2, 1, appearanceTest_time)

        elif reply == QMessageBox.No:
            self.appearance = False
            appearanceEnd_time = time.time()
            appearanceTest_time = round(appearanceEnd_time - appearanceStart_time, 2)
            if not mainThreadRunning():
                return False
            self.itemOperation(mTable, 0, 2, 2, appearanceTest_time)

        self.testNum = self.testNum - 1

        return True

    def labelChange(self,list):
        self.label.setStyleSheet(self.testState_qss[list[0]])
        self.label.setText(list[1])

    def allFinished(self):
        end_time = time.time()
        total_time = round(end_time - self.allStart_time, 3)
        self.showInf(f'本轮测试总时间：{total_time} 秒' + self.HORIZONTAL_LINE)
        # 关闭CAN分析仪
        # CAN_option.close(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX)
        # self.lineEdit_PN.clear()
        self.lineEdit_SN.clear()
        self.lineEdit_REV.clear()
        # self.lineEdit_PN.setReadOnly(False)
        self.lineEdit_SN.setReadOnly(False)
        self.lineEdit_REV.setReadOnly(False)
        if self.tabWidget.currentIndex() == 0:
            # self.lineEdit.clear()
            self.lineEdit_3.clear()
            self.lineEdit_5.clear()
        elif self.tabWidget.currentIndex() == 1:
            # self.lineEdit_22.clear()
            self.lineEdit_21.clear()
            self.lineEdit_20.clear()
        elif self.tabWidget.currentIndex() == 2:
            self.lineEdit_45.clear()
            self.lineEdit_46.clear()
            # self.lineEdit_47.clear()
        elif self.tabWidget.currentIndex() == 3:
            self.lineEdit_31.clear()
            self.lineEdit_32.clear()
            # self.lineEdit_47.clear()
        self.lineEdit_SN.setFocus()
        self.endOfTest()

    def showMessageBox(self,list):
        # 在主线程中创建消息框并显示
        reply = QMessageBox.question(None, list[0], list[1],
                                     QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.Yes)

        # 将弹窗结果放入队列
        self.result_queue.put(reply)

    def PassOrFail(self,isPass):
        if not isPass:
            self.endOfTest()
            self.initPara()
            # self.textBrowser_5.clear()
            self.errorInf = ''
            self.errorNum = 0
            self.isAIPassTest = True
            self.isAIVolPass = True
            self.isAICurPass = True

            self.isAOPassTest = True
            self.isAOVolPass = True
            self.isAOCurPass = True

            self.isLEDRunOK = True
            self.isLEDErrOK = True
            self.CAN_runLED = True
            self.CAN_errorLED = True
            self.stop_subThread()

            # self.label.setStyleSheet(self.testState_qss['stop'])
            # self.label.setText('')
        else:
            self.stop_subThread()

        self.pushButton_10.setEnabled(True)
        self.pushButton_10.setStyleSheet(f"""
                                        QPushButton {{
                                            background-color: #11826C;
                                            color:white
                                        }}
                                        QPushButton:hover {{
                                            background-color: {self.hover_color.name()};
                                            color:#212B36
                                        }}
                                    """)

    def  stop_subThread(self):
        if self.testFlag == 'AI':
            if self.AI_thread and self.AI_thread.isRunning():
                try:
                    self.AI_thread.quit()
                    self.AI_thread.wait()
                    # self.showInf(f'结束AI子线程成功！' + self.HORIZONTAL_LINE)
                except Exception as e:
                    self.showInf(f'结束AI线程异常！' + self.HORIZONTAL_LINE)
                    traceback.print_exc()
        elif self.testFlag == 'AO':
            if self.AO_thread and self.AO_thread.isRunning():
                try:
                    self.AO_thread.quit()
                    self.AO_thread.wait()
                except Exception as e:
                    self.showInf(f'结束AO线程异常！' + self.HORIZONTAL_LINE)
                    traceback.print_exc()
        elif self.testFlag == 'DO' or self.testFlag == 'DI' or self.testFlag == 'DIDO':
            if self.DIDO_thread and self.DIDO_thread.isRunning():
                # self.showInf(f'结束DIDO子线程' + self.HORIZONTAL_LINE)
                try:
                    self.DIDO_thread.quit()
                    self.DIDO_thread.wait()
                except Exception as e:
                    self.showInf(f'结束DIDO线程异常！' + self.HORIZONTAL_LINE)
                    traceback.print_exc()
        elif self.testFlag == 'CPU':
            if self.CPU_thread and self.CPU_thread.isRunning():
                try:
                # self.showInf(f'结束CPU子线程' + self.HORIZONTAL_LINE)
                    self.CPU_thread.quit()
                    self.CPU_thread.wait()
                except Exception as e:
                    self.showInf(f'结束CPU线程异常！' + self.HORIZONTAL_LINE)
                    traceback.print_exc()
        else:
            self.showInf(f'不存在运行的子线程' + self.HORIZONTAL_LINE)
    
        

    def DIDOCANAddr_stateChanged(self):
        self.lineEdit_6.setEnabled(self.checkBox_27.isChecked())
        self.label_16.setEnabled(self.checkBox_27.isChecked())
        self.lineEdit_7.setEnabled(self.checkBox_27.isChecked())
        self.label_17.setEnabled(self.checkBox_27.isChecked())
        # self.saveConfig()

    def DIDOAdPara_stateChanged(self):
        self.lineEdit_13.setEnabled(self.checkBox_28.isChecked())
        self.label_22.setEnabled(self.checkBox_28.isChecked())
        self.lineEdit_12.setEnabled(self.checkBox_28.isChecked())
        self.label_23.setEnabled(self.checkBox_28.isChecked())
        self.lineEdit_14.setEnabled(self.checkBox_28.isChecked())
        self.label_26.setEnabled(self.checkBox_28.isChecked())
        self.label_25.setEnabled(self.checkBox_28.isChecked())
        self.label_24.setEnabled(self.checkBox_28.isChecked())
        # self.saveConfig()
    """

    AI

    """

    def AICANAddr_stateChanged(self):
        self.lineEdit_18.setEnabled(self.checkBox_4.isChecked())
        self.label_48.setEnabled(self.checkBox_4.isChecked())
        self.lineEdit_19.setEnabled(self.checkBox_4.isChecked())
        self.label_49.setEnabled(self.checkBox_4.isChecked())
        self.lineEdit_23.setEnabled(self.checkBox_4.isChecked())
        self.label_54.setEnabled(self.checkBox_4.isChecked())
        # self.saveConfig()

    def AIAdPara_stateChanged(self):
        self.lineEdit_15.setEnabled(self.checkBox_3.isChecked())
        self.label_42.setEnabled(self.checkBox_3.isChecked())
        self.lineEdit_16.setEnabled(self.checkBox_3.isChecked())
        self.label_43.setEnabled(self.checkBox_3.isChecked())
        self.lineEdit_17.setEnabled(self.checkBox_3.isChecked())
        self.label_44.setEnabled(self.checkBox_3.isChecked())
        self.label_45.setEnabled(self.checkBox_3.isChecked())
        self.label_46.setEnabled(self.checkBox_3.isChecked())
        # self.saveConfig()

    def CPU_paramChanged(self):
        CPU_param_array=[ self.lineEdit_34, self.lineEdit_35, self.lineEdit_36,
                         self.lineEdit_37, self.lineEdit_38,self.comboBox_20,self.comboBox_21,
                         self.comboBox_22,self.comboBox_23,self.label_59,self.label_60,
                         self.label_64,self.label_66,self.label_67,self.label_68,self.label_69,
                         self.label_70,self.label_71,self.label_90]
        for Cparam in CPU_param_array:
            Cparam.setEnabled(self.checkBox_71.isChecked())

        # self.saveConfig()

    def testAllorNot(self):
        CPU_test_array = [self.checkBox_50, self.checkBox_51, self.checkBox_52, self.checkBox_53,
                                self.checkBox_55, self.checkBox_56, self.checkBox_57,
                               self.checkBox_58, self.checkBox_59, self.checkBox_60, self.checkBox_61,
                               self.checkBox_62, self.checkBox_63, self.checkBox_64, self.checkBox_65,
                               self.checkBox_66,self.checkBox_54, self.checkBox_67
                               ]
        if self.checkBox_73.isChecked():
            self.checkBox_73.setText("取消全选")
            for i in range(len(CPU_test_array)):
                CPU_test_array[i].setChecked(True)
        else:
            self.checkBox_73.setText("全选")
            for i in range(len(CPU_test_array)):
                CPU_test_array[i].setChecked(False)
    def AI_NotCalibrate(self):
        self.checkBox_5.setEnabled(False)
        self.checkBox_6.setEnabled(False)

    def AI_Calibrate(self):
        self.checkBox_5.setEnabled(True)
        self.checkBox_6.setEnabled(True)

    def AI_NotTest(self):
        self.checkBox_7.setEnabled(False)
        self.checkBox_8.setEnabled(False)
        self.checkBox_9.setEnabled(False)
        self.checkBox_10.setEnabled(False)

    def AI_Test(self):
        self.checkBox_7.setEnabled(True)
        self.checkBox_8.setEnabled(True)
        self.checkBox_9.setEnabled(True)
        self.checkBox_10.setEnabled(True)

    """

    AO

    """

    def AOCANAddr_stateChanged(self):
        self.lineEdit_39.setEnabled(self.checkBox_33.isChecked())
        self.label_74.setEnabled(self.checkBox_33.isChecked())
        self.lineEdit_40.setEnabled(self.checkBox_33.isChecked())
        self.label_75.setEnabled(self.checkBox_33.isChecked())
        self.lineEdit_41.setEnabled(self.checkBox_33.isChecked())
        self.label_76.setEnabled(self.checkBox_33.isChecked())
        # self.saveConfig()

    def AOAdPara_stateChanged(self):
        self.lineEdit_42.setEnabled(self.checkBox_34.isChecked())
        self.label_77.setEnabled(self.checkBox_34.isChecked())
        self.lineEdit_43.setEnabled(self.checkBox_34.isChecked())
        self.label_78.setEnabled(self.checkBox_34.isChecked())
        self.lineEdit_44.setEnabled(self.checkBox_34.isChecked())
        self.label_79.setEnabled(self.checkBox_34.isChecked())
        self.label_80.setEnabled(self.checkBox_34.isChecked())
        self.label_81.setEnabled(self.checkBox_34.isChecked())
        # self.saveConfig()

    def AO_NotCalibrate(self):
        self.checkBox_35.setEnabled(False)
        self.checkBox_36.setEnabled(False)

    def AO_Calibrate(self):
        self.checkBox_35.setEnabled(True)
        self.checkBox_36.setEnabled(True)

    def AO_NotTest(self):
        self.checkBox_29.setEnabled(False)
        self.checkBox_30.setEnabled(False)
        self.checkBox_31.setEnabled(False)
        self.checkBox_32.setEnabled(False)

    def AO_Test(self):
        self.checkBox_29.setEnabled(True)
        self.checkBox_30.setEnabled(True)
        self.checkBox_31.setEnabled(True)
        self.checkBox_32.setEnabled(True)

    def allScreen(self):
        self.label_7.setVisible(False)
        self.isAllScreen = True
        for tW in self.tableWidget_array:
            tW.setVisible(False)
        # self.tableWidget_AI.setVisible(False)
        # self.tableWidget_AO.setVisible(False)
        self.pushbutton_allScreen.setEnabled(False)
        self.pushbutton_allScreen.setVisible(False)
        self.label_28.setGeometry(10, 210, 192, 30)
        self.textBrowser_5.setMaximumSize(711,800)
        self.textBrowser_5.setGeometry(10, 240, 711, 601)

        self.pushbutton_cancelAllScreen.setEnabled(True)
        self.pushbutton_cancelAllScreen.setVisible(True)
        QApplication.processEvents()

    def cancelAllScreen(self):
        self.label_7.setVisible(True)
        self.isAllScreen = False
        self.pushbutton_cancelAllScreen.setEnabled(False)
        self.pushbutton_cancelAllScreen.setVisible(False)

        self.textBrowser_5.setGeometry(10, 664, 711, 177)
        self.label_28.setGeometry(10, 634, 192, 30)
        self.tabIndex = self.tabWidget.currentIndex()

        self.pushbutton_allScreen.setEnabled(True)
        self.pushbutton_allScreen.setVisible(True)
        if self.tabIndex == 0:
            self.tableWidget_DIDO.setVisible(True)
        elif self.tabIndex == 1:
            self.tableWidget_AI.setVisible(True)
        elif self.tabIndex == 2:
            self.tableWidget_AO.setVisible(True)
        elif self.tabIndex == 3:
            self.tableWidget_CPU.setVisible(True)
        QApplication.processEvents()
    
    
    def sendMessage(self):
        self.textBrowser_5.clear()
        #清空ASCII码列表
        self.asciiCode_pn = []
        self.asciiCode_sn = []
        self.asciiCode_rev = []

        self.errorInf = ''
        self.errorNum = 0
        self.isAIPassTest = True
        self.isAIVolPass = True
        self.isAICurPass = True

        self.isAOPassTest = True
        self.isAOVolPass = True
        self.isAOCurPass = True

        # DI测试是否通过
        isDIPassTest = True
        # DO测试是否通过
        isDOPassTest = True

        self.isLEDRunOK = True
        self.isLEDErrOK = True
        self.CAN_runLED = True
        self.CAN_errorLED = True

        # self.label.setStyleSheet(self.testState_qss['stop'])
        # self.label.setText('UNTESTED')


        if self.tabIndex == 0:#DI/DO界面
            mTable = self.tableWidget_DIDO
            # print(f'tabIndex={self.tabIndex}')

            self.testNum = 3#外观检测 + CAN_RunErr检测 + RunErr检测 + 通道检测
            # 获取产品信息
            self.module_type = self.comboBox.currentText()
            if self.module_type == 'ET0800' or self.module_type == 'ET1600' or \
                    self.module_type == 'ET3200':
                self.m_Channels = int(self.module_type[2:4])
            elif self.module_type == 'QN0008' or self.module_type == 'QN0016' or \
                    self.module_type == 'QN0032':
                self.m_Channels = int(self.module_type[4:])

            self.showInf(f'模块类型：{self.module_type}，'
                         f'通道：{self.m_Channels}' + self.HORIZONTAL_LINE)
            self.module_pn = self.lineEdit.text()
            self.module_sn = self.lineEdit_3.text()
            self.module_rev = self.lineEdit_5.text()
            # print(f'{self.module_type},{self.module_pn},{self.module_sn},{self.module_rev}')

            if self.comboBox.currentIndex() == 0 or self.comboBox.currentIndex() == 1 \
                    or self.comboBox.currentIndex() == 2:
                self.module_1 = self.inf['DO1']
                self.module_2 = self.inf['DI2']
                self.m_Channels = int(self.module_type[2:4])
                # self.inf_param = [mTable, self.module_1, self.module_2, self.testNum]
                # 获取CAN地址
                self.CANAddr_DI = int(self.lineEdit_7.text())
                self.CANAddr_DO = int(self.lineEdit_6.text())
                self.CAN1 = self.CANAddr_DO
                self.CAN2 = self.CANAddr_DI

            if self.comboBox.currentIndex() == 3 or self.comboBox.currentIndex() == 4 \
                    or self.comboBox.currentIndex() == 5:
                self.module_1 = self.inf['DI1']
                self.module_2 = self.inf['DO2']
                self.m_Channels = int(self.module_type[4:])
                # self.inf_param = [mTable, self.module_1, self.module_2, self.testNum]

                # 获取CAN地址
                self.CANAddr_DI = int(self.lineEdit_6.text())
                self.CANAddr_DO = int(self.lineEdit_7.text())
                self.CAN1 = self.CANAddr_DI
                self.CAN2 = self.CANAddr_DO
            # 获取附加参数
            self.baud_rate = int(self.lineEdit_13.text())
            self.waiting_time = int(self.lineEdit_12.text())
            self.loop_num = int(self.lineEdit_14.text())
            self.saveDir = self.label_41.text()  # 保存路径
            # # print(f'{self.module_type},{self.module_pn},{self.module_sn},{self.module_rev},{self.CANAddr_AI}')

            # 获取测试信息
            self.isTestCANRunErr = self.checkBox.isChecked()
            self.isTestRunErr = self.checkBox_2.isChecked()

            self.itemOperation(mTable, 0, 3, 0, '')
            self.itemOperation(mTable, 5, 3, 0, '')
            if self.isTestRunErr:
                self.itemOperation(mTable, 1, 3, 0, '')
                self.itemOperation(mTable, 2, 3, 0, '')
            elif not self.isTestRunErr:
                self.itemOperation(mTable, 1, 0, 0, '')
                self.itemOperation(mTable, 2, 0, 0, '')
            if self.isTestCANRunErr:
                self.itemOperation(mTable, 3, 3, 0, '')
                self.itemOperation(mTable, 4, 3, 0, '')
            elif not self.isTestCANRunErr:
                self.itemOperation(mTable, 3, 0, 0, '')
                self.itemOperation(mTable, 4, 0, 0, '')


            self.inf_param = [mTable, self.module_1, self.module_2, self.testNum]
            self.inf_product = [self.module_type, self.module_pn, self.module_sn, self.module_rev, self.m_Channels]
            self.inf_CANAdrr = [self.CANAddr_DI, self.CANAddr_DO]
            self.inf_additional = [self.baud_rate, self.waiting_time, self.loop_num, self.saveDir]
            self.inf_test = [self.isTestCANRunErr, self.isTestRunErr]
            self.inf_DIDOlist = [self.inf_param, self.inf_product, self.inf_CANAdrr,self.inf_additional, self.inf_test]
            # self.textBrowser_5.clear()
            # self.textBrowser_5.insertPlainText(f'产品信息下发成功，可以开始测试。' + self.HORIZONTAL_LINE)
            # self.pushButton_4.setEnabled(True)
            # self.pushButton_4.setStyleSheet("color: rgb(255, 255, 255);\n"
            #                                 "font: bold 25pt \"宋体\";\n"
            #                                 "background-color: rgb(76, 165, 132);")


        elif self.tabIndex == 1:#AI界面
            mTable = self.tableWidget_AI
            # # print(f'tabIndex={self.tabIndex}')
            self.module_1 = self.inf['AO1']
            self.module_2 = self.inf['AI2']
            self.testNum = 4  # CAN_RunErr检测 + RunErr检测 + 电流检测 + 电压检测
            self.inf_param = [mTable, self.module_1,self.module_2,self.testNum]
            # 获取产品信息
            self.module_type = self.comboBox_3.currentText()
            self.m_Channels = int(self.module_type[2:4])
            self.AI_Channels = self.m_Channels
            self.module_pn = self.lineEdit_22.text()
            self.module_sn = self.lineEdit_21.text()
            self.module_rev = self.lineEdit_20.text()

            self.inf_product = [self.module_type,self.module_pn,self.module_sn,self.module_rev,self.m_Channels]
            # # print(f'{self.module_type},{self.module_pn},{self.module_sn},{self.module_rev},{self.m_Channels}')
            # 获取CAN地址
            self.CANAddr_AI = int(self.lineEdit_19.text())
            self.CANAddr_AO = int(self.lineEdit_18.text())
            self.CAN1 = self.CANAddr_AO
            self.CAN2 = self.CANAddr_AI
            self.CANAddr_relay = int(self.lineEdit_23.text())
            # self.CANAddr_relay = 0x203
            self.inf_CANAdrr = [self.CANAddr_AI,self.CANAddr_AO,self.CANAddr_relay]
            # 获取附加参数
            self.baud_rate = int(self.lineEdit_16.text())
            self.waiting_time = int(self.lineEdit_17.text())
            self.receive_num = int(self.lineEdit_15.text())
            self.saveDir = self.label_41.text()  # 保存路径
            # # print(f'{self.module_type},{self.module_pn},{self.module_sn},{self.module_rev},{self.CANAddr_AI}')
            self.inf_additional = [self.baud_rate, self.waiting_time, self.receive_num,self.saveDir]
            # 获取标定信息
            if (self.radioButton.isChecked()):
                self.isCalibrate = False
                # # print(f'isCalibrate:{self.isCalibrate}')
                self.isCalibrateVol = False
                self.isCalibrateCur = False
                self.inf_calibrate = [self.isCalibrate, self.isCalibrateVol, self.isCalibrateCur, 0]
            elif (self.radioButton_2.isChecked() or self.radioButton_3.isChecked()):
                self.isCalibrate = True
                # print(f'isCalibrate:{self.isCalibrate}')
                self.isCalibrateVol = self.checkBox_5.isChecked()
                # print(f'isCalibrateVol:{self.isCalibrateVol}')
                self.isCalibrateCur = self.checkBox_6.isChecked()
                # print(f'isCalibrateCur:{self.isCalibrateCur}')
                if self.radioButton_2.isChecked():
                    self.inf_calibrate = [self.isCalibrate,self.isCalibrateVol,self.isCalibrateCur,1]
                elif self.radioButton_3.isChecked():
                    self.inf_calibrate = [self.isCalibrate, self.isCalibrateVol, self.isCalibrateCur,2]
                # elif not self.radioButton_2.isChecked() and not self.radioButton_3.isChecked():
                #     self.inf_calibrate = [self.isCalibrate, self.isCalibrateVol, self.isCalibrateCur,0]
            else:
                self.inf_calibrate = [self.isCalibrate, self.isCalibrateVol, self.isCalibrateCur, 0]
            # 获取检测信息
            if (self.radioButton_4.isChecked()):
                self.isTest = False
                # # print(f'isCalibrate:{self.isCalibrate}')
                self.isAITestVol = False
                self.isAITestCur = False
                self.isTestCANRunErr = False
                self.isTestRunErr = False
            elif (self.radioButton_5.isChecked()):
                self.isTest = True
                # # print(f'isTest:{self.isTest}')
                self.isAITestVol = self.checkBox_9.isChecked()
                # # print(f'isAITestVol:{self.isAITestVol}')
                self.isAITestCur = self.checkBox_10.isChecked()
                # # print(f'isAITestCur:{self.isAITestCur}')
                self.isTestCANRunErr = self.checkBox_8.isChecked()
                # # print(f'isTestCANRunErr:{self.isTestCANRunErr}')
                self.isTestRunErr = self.checkBox_7.isChecked()
                # # print(f'isTestRunErr:{self.isTestRunErr}')
            else:
                self.isTest = False
                # # print(f'isCalibrate:{self.isCalibrate}')
                self.isAITestVol = False
                self.isAITestCur = False
                self.isTestCANRunErr = False
                self.isTestRunErr = False
            self.inf_test = [self.isTest,self.isAITestVol,self.isAITestCur,self.isTestCANRunErr,self.isTestRunErr]

            self.itemOperation(mTable, 0, 3, 0, '')
            if self.isTestRunErr:
                self.itemOperation(mTable, 1, 3, 0, '')
                self.itemOperation(mTable, 2, 3, 0, '')
            elif not self.isTestRunErr:
                self.itemOperation(mTable, 1, 0, 0, '')
                self.itemOperation(mTable, 2, 0, 0, '')
            if self.isTestCANRunErr:
                self.itemOperation(mTable, 3, 3, 0, '')
                self.itemOperation(mTable, 4, 3, 0, '')
            elif not self.isTestCANRunErr:
                self.itemOperation(mTable, 3, 0, 0, '')
                self.itemOperation(mTable, 4, 0, 0, '')
            if self.isCalibrateVol:
                self.itemOperation(mTable, 5, 3, 0, '')
            elif not self.isCalibrateVol:
                self.itemOperation(mTable, 5, 0, 0, '')
            if self.isCalibrateCur:
                self.itemOperation(mTable, 6, 3, 0, '')
            elif not self.isCalibrateCur:
                self.itemOperation(mTable, 6, 0, 0, '')
            if self.isAITestVol:
                self.itemOperation(mTable, 7, 3, 0, '')
            elif not self.isAITestVol:
                self.itemOperation(mTable, 7, 0, 0, '')
            if self.isAITestCur:
                self.itemOperation(mTable, 8, 3, 0, '')
            elif not self.isAITestCur:
                self.itemOperation(mTable, 8, 0, 0, '')

            self.inf_AIlist = [self.inf_param, self.inf_product, self.inf_CANAdrr,
                               self.inf_additional, self.inf_calibrate, self.inf_test]

        elif self.tabIndex == 2:#AO界面
            mTable = self.tableWidget_AO
            # print(f'tabIndex={self.tabIndex}')
            self.module_1 = self.inf['AI1']
            self.module_2 = self.inf['AO2']
            self.testNum = 4  # CAN_RunErr检测 + RunErr检测 + 电流检测 + 电压检测
            self.inf_param = [mTable, self.module_1, self.module_2, self.testNum]
            #获取产品信息
            self.module_type = self.comboBox_5.currentText()
            self.m_Channels = int(self.module_type[4:])
            self.AO_Channels = self.m_Channels
            self.module_pn = self.lineEdit_47.text()
            self.module_sn = self.lineEdit_46.text()
            self.module_rev = self.lineEdit_45.text()
            # print(f'{self.module_type},{self.module_pn},{self.module_sn},{self.module_rev},{self.m_Channels}')
            self.inf_product = [self.module_type, self.module_pn, self.module_sn, self.module_rev, self.m_Channels]

            #获取CAN地址
            self.CANAddr_AI = int(self.lineEdit_39.text())
            self.CANAddr_AO = int(self.lineEdit_40.text())
            self.CAN1 = self.CANAddr_AI
            self.CAN2 = self.CANAddr_AO
            self.CANAddr_relay = int(self.lineEdit_41.text())
            # self.CANAddr_relay = 0x203
            self.inf_CANAdrr = [self.CANAddr_AI, self.CANAddr_AO, self.CANAddr_relay]
            #获取附加参数
            self.baud_rate = int(self.lineEdit_42.text())
            self.waiting_time = int(self.lineEdit_43.text())
            self.receive_num = int(self.lineEdit_44.text())
            self.saveDir = self.label_41.text()  # 保存路径
            # # print(f'{self.module_type},{self.module_pn},{self.module_sn},{self.module_rev},{self.CANAddr_AI}')
            self.inf_additional = [self.baud_rate, self.waiting_time, self.receive_num, self.saveDir]
            #获取标定信息
            if(self.radioButton_18.isChecked()):
                self.isCalibrate = False
                # # print(f'isCalibrate:{self.isCalibrate}')
                self.isCalibrateVol = False
                self.isCalibrateCur = False
                self.inf_calibrate = [self.isCalibrate, self.isCalibrateVol, self.isCalibrateCur, 0]
            elif(self.radioButton_19.isChecked() or self.radioButton_20.isChecked()):
                self.isCalibrate = True
                # print(f'isCalibrate:{self.isCalibrate}')
                self.isCalibrateVol = self.checkBox_35.isChecked()
                # print(f'isCalibrateVol:{self.isCalibrateVol}')
                self.isCalibrateCur = self.checkBox_36.isChecked()
                # print(f'isCalibrateCur:{self.isCalibrateCur}')
                if self.radioButton_19.isChecked():
                    self.inf_calibrate = [self.isCalibrate,self.isCalibrateVol,self.isCalibrateCur,1]
                elif self.radioButton_20.isChecked():
                    self.inf_calibrate = [self.isCalibrate, self.isCalibrateVol, self.isCalibrateCur,2]
            else:
                self.inf_calibrate = [self.isCalibrate, self.isCalibrateVol, self.isCalibrateCur, 0]
            # 获取检测信息
            if (self.radioButton_16.isChecked()):
                self.isTest = False
                # # print(f'isCalibrate:{self.isCalibrate}')
                self.isAOTestVol = False
                self.isAOTestCur = False
                self.isTestCANRunErr = False
                self.isTestRunErr = False
            elif (self.radioButton_17.isChecked()):
                self.isTest = True
                # # print(f'isTest:{self.isTest}')
                self.isAOTestVol = self.checkBox_29.isChecked()
                # # print(f'isAOTestVol:{self.isAOTestVol}')
                self.isAOTestCur = self.checkBox_30.isChecked()
                # # print(f'isAOTestCur:{self.isAOTestCur}')
                self.isTestCANRunErr = self.checkBox_31.isChecked()
                # # print(f'isTestCANRunErr:{self.isTestCANRunErr}')
                self.isTestRunErr = self.checkBox_32.isChecked()
                # # print(f'isTestRunErr:{self.isTestRunErr}')
            else:
                self.isTest = False
                # # print(f'isCalibrate:{self.isCalibrate}')
                self.isAOTestVol = False
                self.isAOTestCur = False
                self.isTestCANRunErr = False
                self.isTestRunErr = False
            self.inf_test = [self.isTest, self.isAOTestVol, self.isAOTestCur, self.isTestCANRunErr, self.isTestRunErr]

            self.itemOperation(mTable, 0, 3, 0, '')
            if self.isTestRunErr:
                self.itemOperation(mTable, 1, 3, 0, '')
                self.itemOperation(mTable, 2, 3, 0, '')
            elif not self.isTestRunErr:
                self.itemOperation(mTable, 1, 0, 0, '')
                self.itemOperation(mTable, 2, 0, 0, '')
            if self.isTestCANRunErr:
                self.itemOperation(mTable, 3, 3, 0, '')
                self.itemOperation(mTable, 4, 3, 0, '')
            elif not self.isTestCANRunErr:
                self.itemOperation(mTable, 3, 0, 0, '')
                self.itemOperation(mTable, 4, 0, 0, '')
            if self.isCalibrateVol:
                self.itemOperation(mTable, 5, 3, 0, '')
            elif not self.isCalibrateVol:
                self.itemOperation(mTable, 5, 0, 0, '')
            if self.isCalibrateCur:
                self.itemOperation(mTable, 6, 3, 0, '')
            elif not self.isCalibrateCur:
                self.itemOperation(mTable, 6, 0, 0, '')
            if self.isAOTestVol:
                self.itemOperation(mTable, 7, 3, 0, '')
            elif not self.isAOTestVol:
                self.itemOperation(mTable, 7, 0, 0, '')
            if self.isAOTestCur:
                self.itemOperation(mTable, 8, 3, 0, '')
            elif not self.isAOTestCur:
                self.itemOperation(mTable, 8, 0, 0, '')

            self.inf_AOlist = [self.inf_param, self.inf_product, self.inf_CANAdrr,
                               self.inf_additional, self.inf_calibrate, self.inf_test]


        elif self.tabIndex == 3:#CPU界面
            mTable = self.tableWidget_CPU
            # print(f'tabIndex={self.tabIndex}')
            self.module_1 = '工装1（ET1600）'
            self.module_2 = '工装2（QN0016）'
            self.module_3 = '工装3（QN0016）'
            self.module_4 = '工装4（AE0400）'
            self.module_5 = '工装5（AQ0004）'
            self.testNum = 18  # ["外观检测", "型号检查", "SRAM", "FLASH", "FPGA", "拨杆测试", "MFK按键",
                                  #  "掉电保存", "RTC测试","各指示灯", "本体IN", "本体OUT", "以太网",
                                    # "RS-232C", "RS-485","右扩CAN", "MAC/三码写入", "MA0202"]

            self.inf_param = [mTable, self.module_1, self.module_2, self.module_3,
                              self.module_4,self.module_5,self.testNum]
            #获取产品信息
            self.module_type = self.comboBox_20.currentText()
            self.in_Channels = int(self.module_type[5:7])
            self.out_Channels = int(self.module_type[7:9])
            self.module_pn = self.lineEdit_30.text()
            self.module_sn = self.lineEdit_31.text()
            self.module_rev = self.lineEdit_32.text()
            self.module_MAC = self.lineEdit_33.text()
            # print(f'{self.module_type},{self.module_pn},{self.module_sn},{self.module_rev},{self.m_Channels}')
            self.inf_product = [self.module_type, self.module_pn, self.module_sn,
                                self.module_rev, self.module_MAC, self.in_Channels,self.out_Channels]

            #获取CAN与IP地址
            self.CANAddr1 = int(self.lineEdit_34.text())
            self.CANAddr2 = int(self.lineEdit_35.text())
            self.CANAddr3 = int(self.lineEdit_36.text())
            self.CANAddr4 = int(self.lineEdit_37.text())
            self.CANAddr5 = int(self.lineEdit_38.text())

            # self.IPAddr = '192.168.1.66'
            self.inf_CANIPAdrr = [self.CANAddr1,self.CANAddr2,self.CANAddr3,self.CANAddr4,
                                  self.CANAddr5]
            #获取串口信息
            self.serialPort_232 = self.comboBox_21.currentText()
            self.serialPort_485 = self.comboBox_22.currentText()
            self.serialPort_typeC = self.comboBox_23.currentText()
            self.saveDir = self.label_41.text()  # 保存路径
            # # print(f'{self.module_type},{self.module_pn},{self.module_sn},{self.module_rev},{self.CANAddr_AI}')
            self.inf_serialPort = [self.serialPort_232, self.serialPort_485,
                                   self.serialPort_typeC, self.saveDir]
            #获取检测信息

            self.CPU_test_array = [self.checkBox_50, self.checkBox_51, self.checkBox_52,self.checkBox_53,
                                    self.checkBox_55, self.checkBox_56, self.checkBox_57,
                                   self.checkBox_58, self.checkBox_59, self.checkBox_60, self.checkBox_61,
                                   self.checkBox_62, self.checkBox_63, self.checkBox_64, self.checkBox_65,
                                   self.checkBox_66, self.checkBox_54, self.checkBox_67
                                   ]
            self.CPU_testName_array = [ "外观检测", "型号检查", "SRAM", "FLASH", "FPGA",
                                        "拨杆测试","MFK按键","掉电保存", "RTC测试", "各指示灯", "本体IN", "本体OUT",
                                        "以太网","RS-232C","RS-485","右扩CAN", "MAC/三码写入", "MA0202"]
            self.inf_CPU_test = [False for x in range(len(self.CPU_test_array))]
            for i in range(len(self.CPU_test_array)):
                self.inf_CPU_test[i] = self.CPU_test_array[i].isChecked()
                # print(f' self.CPU_test_array[{i}].isChecked():{ self.CPU_test_array[i].isChecked()}')
            self.inf_test = [self.inf_CPU_test,self.checkBox_68.isChecked()]
            # :param
            # mTable: 进行操作的表格
            # :param
            # row: 进行操作的单元行
            # :param
            # state: 测试状态['无需检测', '正在检测', '检测完成', '等待检测']
            # :param
            # result: 测试结果
            # col2 = ['', '通过', '未通过']
            # :param
            # operationTime: 测试时间
            for i in range(len(self.CPU_test_array)):
                if self.inf_CPU_test[i]:
                    self.itemOperation(mTable, i, 3, 0, '')
                else:
                    self.itemOperation(mTable, i, 0, 0, '')
            self.inf_CPUlist = [self.inf_param,self.inf_product, self.inf_CANIPAdrr,
                                self.inf_serialPort, self.inf_test]
        if self.tabIndex == 0 or self.tabIndex == 1 or self.tabIndex == 2:
            #三码转换
            self.asciiCode_pn = (strToASCII(self.lineEdit_PN.text()))
            self.asciiCode_sn = (strToASCII(self.lineEdit_SN.text()))
            self.asciiCode_rev = (strToASCII(self.lineEdit_REV.text()))

            #三码写入PLC
            if not write3codeToPLC(self.CAN2,[self.asciiCode_pn,self.asciiCode_sn,self.asciiCode_rev]):
                return False
            else:
                self.showInf('三码写入成功！\n\n')

        self.textBrowser_5.clear()
        self.textBrowser_5.insertPlainText(f'产品信息下发成功，开始测试。' + self.HORIZONTAL_LINE)


        self.pushButton_pause.setEnabled(True)
        self.pushButton_pause.setVisible(True)

        self.pushButton_resume.setEnabled(False)
        self.pushButton_resume.setVisible(False)
        return True
        # self.pushButton_4.setStyleSheet(self.topButton_qss['on'])
        # self.pushButton_4.setEnabled(True)
        # self.pushButton_4.setVisible(True)

    def clearList(self,array):
        for i in range(len(array)):
            array[i] = 0x00

    #自动分配节点
    def configCANAddr(self,addr1,addr2,addr3,addr4,addr5):
        if self.tabIndex == 1 or self.tabIndex == 2:
            list =[addr1,addr2,addr3,addr4]
        elif self.tabIndex == 3:
            list = [addr1, addr2, addr3, addr4, addr5]
        else:
            list = [addr1, addr2]
        for a in list:
            self.m_transmitData=[0xac, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
            self.m_transmitData[2] = a
            boola = CAN_option.transmitCANAddr(0x0,self.m_transmitData)[0]
            time.sleep(0.01)
            if not boola:
                self.showInf(f'节点{a}分配失败' + self.HORIZONTAL_LINE)
                return False
        self.showInf(f'所有节点分配成功' + self.HORIZONTAL_LINE)
        return True


    def isModulesOnline(self):
        # 检测设备心跳
        if self.check_heartbeat(self.CAN1, self.module_1, self.waiting_time) == False:
            return False
        if self.check_heartbeat(self.CAN2, self.module_2, self.waiting_time) == False:
            return False
        if self.tabIndex !=0:
            if self.check_heartbeat(self.CANAddr_relay, '继电器1', self.waiting_time) == False:
                return False
            if self.check_heartbeat(self.CANAddr_relay+1, '继电器2', self.waiting_time) == False:
                return False

        return True

    def check_heartbeat(self, can_addr, inf, max_waiting):
        if inf == '继电器':
            bool_receive, self.m_can_obj = CAN_option.receiveCANbyID(0x700 + can_addr , max_waiting,0)
            # print(self.m_can_obj.Data)
            if bool_receive == False:
                self.showInf(f'错误：未发现{inf}' + self.HORIZONTAL_LINE)
                # self.isPause()
                # if not self.isStop():
                #     return
                return False

            self.showInf(f'发现{inf}：收到心跳帧：{hex(self.m_can_obj.ID)}\n\n')


        else:
            can_id = 0x700 + can_addr
            bool_receive, self.m_can_obj = CAN_option.receiveCANbyID(can_id, max_waiting,0)
            # print(self.m_can_obj.Data)
            if bool_receive == False:
                self.showInf(f'错误：未发现{inf}' + self.HORIZONTAL_LINE)
                # self.isPause()
                # if not self.isStop():
                #     return
                return False

            self.showInf(f'发现{inf}：收到心跳帧：{hex(self.m_can_obj.ID)}\n\n')
        # self.isPause()
        #if not self.isStop():
#            return
        return True

    def getSerialInf(self, num:int):
        self.textBrowser_5.clear()
        #清空串口选项
        for i in range(self.comboBox_21.count()):
            self.comboBox_21.removeItem(i)
            self.comboBox_22.removeItem(i)
            self.comboBox_23.removeItem(i)
        # 获取电脑上可用的串口列表
        ports = serial.tools.list_ports.comports()

        # 遍历串口列表并打印串口信息
        for i in range(len(ports)):
            self.comboBox_21.addItem("")
            self.comboBox_22.addItem("")
            self.comboBox_23.addItem("")
            self.comboBox_21.setItemText(i, ports[i].device)
            self.comboBox_22.setItemText(i, ports[i].device)
            self.comboBox_23.setItemText(i, ports[i].device)
            if num != 0:
                if i ==0:
                    self.showInf(f'可用串口：\n')
                self.showInf(f'({i + 1}){ports[i].description}\n')
            # self.showInf(f'({i+1}){ports[i].device}, {ports[i].name}, {ports[i].description}\n')

        self.comboBox_21.removeItem(len(ports))
        self.comboBox_22.removeItem(len(ports))
        self.comboBox_23.removeItem(len(ports))
    def CANFail(self):
        self.endOfTest()
        self.initPara()
        # self.textBrowser_5.clear()
        self.errorInf = ''
        self.errorNum = 0
        self.isAIPassTest = True
        self.isAIVolPass = True
        self.isAICurPass = True

        self.isAOPassTest = True
        self.isAOVolPass = True
        self.isAOCurPass = True

        self.isLEDRunOK = True
        self.isLEDErrOK = True
        self.CAN_runLED = True
        self.CAN_errorLED = True
        self.reInputPNSNREV()

    def saveExcel(self,saveList):
        saveList[0].save(str(self.label_41.text()) + saveList[1])
        # book.save(self.label_41.text() + saveDir)

    def printResult(self,list):
        try:
            self.generateLabel(list)
            content = list[1]
            file_name = f'{self.label_41.text()}{list[0]}_label.docx'
            # if list[1] == 'FAIL':
            #     content += f'\n\n{list[2]}'
            #
            # with open(file_name, "w") as f:
            #     for line in content.splitlines():
            #         f.write(line + "\n")

            win32api.ShellExecute(
                0,
                "print",
                file_name,
                #
                # If this is None, the default printer will
                # be used anyway.
                #
                '/d:"%s"' % win32print.GetDefaultPrinter(),
                ".",
                0
            )
        except Exception as e:
            self.showInf(f"printResultError:{e}+{self.HORIZONTAL_LINE}")
            traceback.print_exc()
    def generateLabel(self,list):
        from docx import Document

        # docx.shared 用于设置大小（图片等）

        from docx.shared import Cm, Pt
        from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
        from docx.document import Document as Doc
        try:
            # 创建代表Word文档的Doc对象

            document = Document()
            default_section = document.sections[0]
            # 默认宽度和高度
            # # print(default_section.page_width.cm)  # 21.59
            # # print(default_section.page_height.cm)  # 27.94
            # 可直接修改宽度和高度，即纸张大小改为自定义
            default_section.page_width = Cm(8)
            default_section.page_height = Cm(12)

            # 修改页边距
            default_section.top_margin = Cm(0.5)
            default_section.right_margin = Cm(0.5)
            default_section.bottom_margin = Cm(0.5)
            default_section.left_margin = Cm(0.5)

            # 添加图片（注意路径和图片必须要存在）
            document.add_picture(self.current_dir+'/logo.png', width=Cm(6.1))

            # # 添加带样式的段落
            p = document.add_paragraph('')
            p.paragraph_format.line_spacing = 1
            PLC = p.add_run('PLC I/O')
            PLC.font.name = 'Stencil'
            PLC.font.size = Pt(24)
            PLC.underline = False

            s = p.add_run('s fac_Test')
            s.font.name = 'Stencil'
            s.font.size = Pt(16)
            s.underline = False

            pSN = document.add_paragraph('')
            pSN.paragraph_format.line_spacing = 1
            SN = pSN.add_run(f'UUT SN: {self.module_sn}')
            SN.font.name = '等线'
            SN.font.size = Pt(10.5)

            pPN = document.add_paragraph('')
            pPN.paragraph_format.line_spacing = 1
            PN = pPN.add_run(f'UUT PN: {self.module_pn}')
            PN.font.name = '等线'
            PN.font.size = Pt(10.5)

            pREV = document.add_paragraph('')
            pREV.paragraph_format.line_spacing = 1
            REV = pREV.add_run(f'UUT REV: {self.module_rev}')
            REV.font.name = '等线'
            REV.font.size = Pt(10.5)

            pTime = document.add_paragraph('')
            pTime.paragraph_format.line_spacing = 1
            Time = pTime.add_run(f'Time Duration: {round(time.time() - self.allStart_time, 1)} 秒')
            Time.font.name = '等线'
            Time.font.size = Pt(12)
            Time.bold = True
            if list[1] == 'FAIL':
                # 添加图片（注意路径和图片必须要存在）
                document.add_picture(self.current_dir+'/fail.png', width=Cm(5.5))
                pFail = document.add_paragraph('')
                pFail.paragraph_format.line_spacing = 1
                Fail = pFail.add_run('FAILED ITEMS：')
                Fail.font.name = '等线'
                Fail.font.size = Pt(12)
                Fail.bold = True

                pFail_inf = document.add_paragraph('')
                pFail_inf.paragraph_format.line_spacing = 1
                fail_inf = pFail_inf.add_run(f'{list[2][1:]}')
                fail_inf.font.name = '等线'
                fail_inf.font.size = Pt(10)
            elif list[1] == 'PASS':
                # 添加图片（注意路径和图片必须要存在）
                document.add_picture(self.current_dir+'/pass.png', width=Cm(5.5))
            document.paragraphs[6].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


            # 保存文档
            for p in document.paragraphs:
                p.paragraph_format.line_spacing = 1
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
            document.save(self.label_41.text()+f'{list[0]}_label.docx')
        except Exception as e:
            self.showInf(f"generateLabelError:{e}+{self.HORIZONTAL_LINE}")
            # 捕获异常并输出详细的错误信息
            traceback.print_exc()


    @abstractmethod
    def calibrateAO(self):
        raise NotImplementedError()

    @abstractmethod
    def testAO(self):
        raise NotImplementedError()

    @abstractmethod
    def calibrateAI(self):
        raise NotImplementedError()

    @abstractmethod
    def testAI(self):
        raise NotImplementedError()

    @abstractmethod
    def testDI(self):
        raise NotImplementedError()

    @abstractmethod
    def testDO(self):
        raise NotImplementedError()

  #   #检查配套AI模块输入类型是否正确
  #   def setAIInputType(self,AIChannel,type):
  #       #CAN_option.close(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX)
  #       #self.can_start()
  #       if AIChannel == 1:
  #           self.showInf(f'设置AI通道量程：\n\n')
  #
  #       self.m_transmitData[0] = 0x2b
  #       self.m_transmitData[1] = 0x10
  #       self.m_transmitData[2] = 0x61
  #       self.m_transmitData[3] = AIChannel
  #       self.m_transmitData[6] = 0x00
  #       self.m_transmitData[7] = 0x00
  #
  #       self.clearList(self.m_can_obj.Data)
  #
  #       if type == 'AOVoltage' or type == 'AIVoltage':
  #           self.m_transmitData[4] = self.AIRangeArray[0]
  #           self.m_transmitData[5] = self.AIRangeArray[1]
  #           while True:
  #               self.isPause()
  #               #if not self.isStop():
  # #                  return
  #               QApplication.processEvents()
  #               if CAN_option.transmitCAN(0x600 + self.CANAddr_AI, self.m_transmitData)[0]:
  #                   bool_receive, self.m_can_obj = CAN_option.receiveCANbyID(0x580 + self.CANAddr_AI,
  #                                                                            self.waiting_time)
  #                   if bool_receive:
  #                       # # print(f'm_can_obj.Data[4]:{self.m_can_obj.Data[4]}')
  #                       # # print(f'm_can_obj.Data[5]:{self.m_can_obj.Data[5]}')
  #                       # # print(f'AIRangeArray[0]:{self.AIRangeArray[0]}')
  #                       # # print(f'AIRangeArray[1]:{self.AIRangeArray[1]}')
  #                       if hex(self.m_can_obj.Data[4]) == hex(self.AIRangeArray[0]) and hex(
  #                               self.m_can_obj.Data[5]) == hex(self.AIRangeArray[1]):
  #                           # print(f'{AIChannel}.成功设置AI通道{AIChannel}的量程为”电压+/-10V“。\n\n')
  #                           self.showInf(f'{AIChannel}.成功设置AI通道{AIChannel}的量程为”电压+/-10V“。\n\n')
  #                           break
  #                       else:
  #                           self.showInf(f'{AIChannel}.未成功设置AI通道{AIChannel}的量程为”电压+/-10V“。\n\n')
  #                           # print(f'{AIChannel}.未成功设置AI通道{AIChannel}的量程为”电压+/-10V“。\n\n')
  #               else:
  #                   self.showInf(f'{AIChannel}.未成功设置AI通道{AIChannel}的量程为”电压+/-10V“。\n\n')
  #                   # # print(f'{AIChannel}.未成功设置AI通道{AIChannel}的量程为”电压+/-10V“。\n\n')
  #       elif type == 'AOCurrent' or type == 'AICurrent':
  #           self.m_transmitData[4] = self.AIRangeArray[10]
  #           self.m_transmitData[5] = self.AIRangeArray[11]
  #           while True:
  #               self.isPause()
  #               #if not self.isStop():
  # #                  return
  #               QApplication.processEvents()
  #               if CAN_option.transmitCAN(0x600 + self.CANAddr_AI, self.m_transmitData)[0]:
  #                   bool_receive, self.m_can_obj = CAN_option.receiveCANbyID(0x580 + self.CANAddr_AI,
  #                                                                            self.waiting_time)
  #                   if bool_receive:
  #                       # # print(f'm_can_obj.Data[4]:{self.m_can_obj.Data[4]}')
  #                       # # print(f'm_can_obj.Data[5]:{self.m_can_obj.Data[5]}')
  #                       # # print(f'AIRangeArray[0]:{self.AIRangeArray[0]}')
  #                       # # print(f'AIRangeArray[1]:{self.AIRangeArray[1]}')
  #                       if hex(self.m_can_obj.Data[4]) == hex(self.AIRangeArray[10]) and hex(
  #                               self.m_can_obj.Data[5]) == hex(self.AIRangeArray[11]):
  #                           # print(f'{AIChannel}.成功设置AI通道{AIChannel}的量程为”电流(0-20)mA“。\n\n')
  #                           self.showInf(f'{AIChannel}.成功设置AI通道{AIChannel}的量程为”电流(0-20)mA“。\n\n')
  #                           break
  #                       else:
  #                           self.showInf(f'{AIChannel}.未成功设置AI通道{AIChannel}的量程为”电流(0-20)mA“。\n\n')
  #                           # print(f'{AIChannel}.未成功设置AI通道{AIChannel}的量程为”电流(0-20)mA“。\n\n')
  #               else:
  #                   self.showInf(f'{AIChannel}.未成功设置AI通道{AIChannel}的量程为”电流(0-20)mA“。\n\n')
  #                   break
  #
  #       if AIChannel == 4:
  #           self.showInf(self.HORIZONTAL_LINE)
  #
  #
  #   # 检查待检AO模块输入类型是否正确
  #   def setAOInputType(self, AOChannel, type):
  #       # self.m_transmitData[0] = 0x2b
  #       # self.m_transmitData[1] = 0x10
  #       # self.m_transmitData[2] = 0x63
  #       # self.m_transmitData[3] = AOChannel
  #       # self.m_transmitData[4] = 0x00
  #       # self.m_transmitData[5] = 0x00
  #       # self.m_transmitData[6] = 0x00
  #       # self.m_transmitData[7] = 0x00
  #       #
  #       # bool_transmit, self.m_can_obj = CAN_option.transmitCAN((0x600 + self.CANAddr_AO), self.m_transmitData)
  #       #CAN_option.close(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX)
  #       #self.can_start()
  #       if AOChannel == 1:
  #           self.showInf(f'设置AO通道量程：\n\n')
  #
  #       self.m_transmitData[0] = 0x2b
  #       self.m_transmitData[1] = 0x10
  #       self.m_transmitData[2] = 0x63
  #       self.m_transmitData[3] = AOChannel
  #       self.m_transmitData[6] = 0x00
  #       self.m_transmitData[7] = 0x00
  #
  #       self.clearList(self.m_can_obj.Data)
  #
  #       if type == 'AOVoltage' or type == 'AIVoltage':
  #           self.m_transmitData[4] = self.AORangeArray[0]
  #           self.m_transmitData[5] = self.AORangeArray[1]
  #           while True:
  #               self.isPause()
  #               #if not self.isStop():
  # #                  return
  #               QApplication.processEvents()
  #               if CAN_option.transmitCAN(0x600 + self.CANAddr_AO, self.m_transmitData)[0]:
  #                   bool_receive, self.m_can_obj = CAN_option.receiveCANbyID(0x580 + self.CANAddr_AO,
  #                                                                            self.waiting_time)
  #                   if bool_receive:
  #                       # # print(f'm_can_obj.Data[4]:{self.m_can_obj.Data[4]}')
  #                       # # print(f'm_can_obj.Data[5]:{self.m_can_obj.Data[5]}')
  #                       # # print(f'AORangeArray[0]:{self.AORangeArray[0]}')
  #                       # # print(f'AORangeArray[1]:{self.AORangeArray[1]}')
  #                       if hex(self.m_can_obj.Data[4]) == hex(self.AORangeArray[0]) and hex(
  #                               self.m_can_obj.Data[5]) == hex(self.AORangeArray[1]):
  #                           # print(f'{AOChannel}.成功设置AO通道{AOChannel}的量程为”电压+/-10V“。\n\n')
  #                           self.showInf(f'{AOChannel}.成功设置AO通道{AOChannel}的量程为”电压+/-10V“。\n\n')
  #                           break
  #                       else:
  #                           self.showInf(f'{AOChannel}.未成功设置AO通道{AOChannel}的量程为”电压+/-10V“。\n\n')
  #                           # print(f'{AOChannel}.未成功设置AO通道{AOChannel}的量程为”电压+/-10V“。\n\n')
  #               else:
  #                   self.showInf(f'{AOChannel}.未成功设置AO通道{AOChannel}的量程为”电压+/-10V“。\n\n')
  #                   # # print(f'{AOChannel}.未成功设置AO通道{AOChannel}的量程为”电压+/-10V“。\n\n')
  #       elif type == 'AOCurrent' or type == 'AICurrent':
  #           self.m_transmitData[4] = self.AORangeArray[10]
  #           self.m_transmitData[5] = self.AORangeArray[11]
  #           while True:
  #               self.isPause()
  #               #if not self.isStop():
  # #                  return
  #               QApplication.processEvents()
  #               if CAN_option.transmitCAN(0x600 + self.CANAddr_AO, self.m_transmitData)[0]:
  #                   bool_receive, self.m_can_obj = CAN_option.receiveCANbyID(0x580 + self.CANAddr_AO,
  #                                                                            self.waiting_time)
  #                   if bool_receive:
  #                       # # print(f'm_can_obj.Data[4]:{self.m_can_obj.Data[4]}')
  #                       # # print(f'm_can_obj.Data[5]:{self.m_can_obj.Data[5]}')
  #                       # # print(f'AORangeArray[0]:{self.AORangeArray[0]}')
  #                       # # print(f'AORangeArray[1]:{self.AORangeArray[1]}')
  #                       if hex(self.m_can_obj.Data[4]) == hex(self.AORangeArray[10]) and hex(
  #                               self.m_can_obj.Data[5]) == hex(self.AORangeArray[11]):
  #                           # print(f'{AOChannel}.成功设置AO通道{AOChannel}的量程为”电流(0-20)mA“。\n\n')
  #                           self.showInf(f'{AOChannel}.成功设置AO通道{AOChannel}的量程为”电流(0-20)mA“。\n\n')
  #                           break
  #                       else:
  #                           self.showInf(f'{AOChannel}.未成功设置AO通道{AOChannel}的量程为”电流(0-20)mA“。\n\n')
  #                           # print(f'{AOChannel}.未成功设置AO通道{AOChannel}的量程为”电流(0-20)mA“。\n\n')
  #               else:
  #                   self.showInf(f'{AOChannel}.未成功设置AO通道{AOChannel}的量程为”电流(0-20)mA“。\n\n')
  #                   break
  #
  #       if AOChannel == 4:
  #           self.showInf(self.HORIZONTAL_LINE)
  #
  #   def receiveAIData(self):
  #       can_id = 0x280 + self.CANAddr_AI
  #       time1 = time.time()
  #       while True:
  #           if (time.time() - time1)*1000 > self.waiting_time:
  #               break
  #           self.isPause()
  #           # if not self.isStop():
  #           #     return
  #           bool_receive,self.m_can_obj = CAN_option.receiveCANbyID(can_id, self.waiting_time)
  #           QApplication.processEvents()
  #           if bool_receive:
  #               break
  #       recv = [0,0,0,0]
  #       for i in range(self.m_Channels):
  #           # # print(f'i= {i}')
  #           recv[i] = self.m_can_obj.Data[i*2] | self.m_can_obj.Data[i*2+1] << 8
  #           # # print(f'recv[{i}]={recv[i]}')
  #           self.isPause()
  #           # if not self.isStop():
  #           #     return
  #       return recv
  #
  #   def calibrate_receiveAIData(self,channelNum):
  #       #CAN_option.close(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX)
  #       #self.can_start()
  #       recv = [0, 0, 0, 0]
  #       if channelNum == 1:
  #           recv=[0]
  #       can_id = 0x600 + self.CANAddr_AI
  #       for i in range(channelNum):
  #           self.m_transmitData[0] = 0x40
  #           self.m_transmitData[1] = 0x3c
  #           self.m_transmitData[2] = 0x64
  #           self.m_transmitData[3] = i+1
  #           self.m_transmitData[4] = 0x00
  #           self.m_transmitData[5] = 0x00
  #           self.m_transmitData[6] = 0x00
  #           self.m_transmitData[7] = 0x00
  #           CAN_option.transmitCAN((0x600 + self.CANAddr_AI),self.m_transmitData)
  #           self.isPause()
  #           # if not self.isStop():
  #           #     return
  #           while True:
  #               self.isPause()
  #               #if not self.isStop():
  # #                  return
  #               bool_receive,self.m_can_obj = CAN_option.receiveCANbyID( 0x580 + self.CANAddr_AI, self.waiting_time)
  #               QApplication.processEvents()
  #               if bool_receive:
  #                   break
  #           recv[i] = ((self.m_can_obj.Data[7] << 24|self.m_can_obj.Data[6] << 16)|self.m_can_obj.Data[5] << 8) | \
  #                     self.m_can_obj.Data[4]
  #
  #       return recv

    def showInf(self,inf):
        self.textBrowser_5.append(inf)
        if inf[:8] =='本轮测试总时间：':
            self.move_to_end()
        QApplication.processEvents()
        time.sleep(0.1)

        # self.isPause()
        # if self.work_thread.stopFlag.isSet():
        #     return


    def move_to_end(self):
        self.textBrowser_5.moveCursor(self.textBrowser_5.textCursor().End)

    # def isPause(self):
    #     # # print(f'self.work_thread.pauseFlag:{self.work_thread.pauseFlag.isSet()}')
    #     if self.work_thread.pauseFlag.isSet():
    #         while True:
    #             # # print(f'while:{self.pause_num}')
    #             if self.pause_num == 1 and not self.work_thread.stopFlag.isSet():
    #                 self.pause_num += 1
    #                 self.showInf(self.HORIZONTAL_LINE + '暂停中…………'+self.HORIZONTAL_LINE)
    #
    #             QApplication.processEvents()
    #             # # print('暂停中…………')
    #             if not self.work_thread.pauseFlag.isSet():
    #                 self.pause_num = 1
    #                 break

    def isStop(self):
        if self.work_thread.stopFlag.isSet():
            return False
        return True

    def endOfTest(self):
        self.pushButton_pause.setEnabled(False)
        self.pushButton_pause.setVisible(False)

        self.pushButton_resume.setEnabled(False)
        self.pushButton_resume.setVisible(False)

        self.pushButton_4.setStyleSheet(self.topButton_qss['off'])
        self.pushButton_4.setEnabled(False)
        self.pushButton_4.setVisible(True)

        self.pushButton_3.setStyleSheet(self.topButton_qss['off'])
        self.pushButton_3.setEnabled(False)
        self.reInputPNSNREV()
        # self.pushButton_9.setEnabled(False)
        # self.pushButton_9.setStyleSheet(self.topButton_qss['off'])

    def AI_itemOperation(self,list):
        '''
        :param list = [row,state,result,operationTime]
        :param row: 进行操作的单元行
        :param state: 测试状态 ['无需检测','正在检测','检测完成','等待检测']
        :param result: 测试结果 col2 = ['','通过','未通过']
        :param operationTime: 测试时间
        :return:
        '''
        mTable = self.tableWidget_AI
        col1 = ['无需检测','正在检测','检测完成','等待检测']
        col2 = ['','通过','未通过']
        font = QtGui.QColor(0, 0, 0)
        if list[1] == 0:
            color = QtGui.QColor(255, 255, 255)
            font = QtGui.QColor(197, 197, 197)
        elif list[1] == 1:
            color = QtGui.QColor(255,255,0)
        elif list[1] == 2 and list[2] == 1:
            color = QtGui.QColor(0, 255, 0)
        elif list[1] == 2 and list[2] == 2:
            color = QtGui.QColor(255, 0, 0)
            font = QtGui.QColor(255, 255, 255)
        elif list[1] == 3:
            color = QtGui.QColor(197, 197, 197)

        for i in range(4):
            item = mTable.item(list[0], i)
            item.setBackground(color)
            item.setForeground(font)
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            if i == 1:
                item.setText(col1[list[1]])
            if i == 2:
                item.setText(col2[list[2]])
            if i == 3 and list[1] == 2:
                item.setText(f'{list[3]}s')
            if (i == 3 and list[1] == 0) or (i == 3 and list[1] == 3):
                item.setText(f'{list[3]}')
        QApplication.processEvents()

    def AO_itemOperation(self,list):
        '''
        :param list = [row,state,result,operationTime]
        :param row: 进行操作的单元行
        :param state: 测试状态 ['无需检测','正在检测','检测完成','等待检测']
        :param result: 测试结果 col2 = ['','通过','未通过']
        :param operationTime: 测试时间
        :return:
        '''
        mTable = self.tableWidget_AO
        col1 = ['无需检测','正在检测','检测完成','等待检测']
        col2 = ['','通过','未通过']
        font = QtGui.QColor(0, 0, 0)
        if list[1] == 0:
            color = QtGui.QColor(255, 255, 255)
            font = QtGui.QColor(197, 197, 197)
        elif list[1] == 1:
            color = QtGui.QColor(255,255,0)
        elif list[1] == 2 and list[2] == 1:
            color = QtGui.QColor(0, 255, 0)
        elif list[1] == 2 and list[2] == 2:
            color = QtGui.QColor(255, 0, 0)
            font = QtGui.QColor(255, 255, 255)
        elif list[1] == 3:
            color = QtGui.QColor(197, 197, 197)

        for i in range(4):
            item = mTable.item(list[0], i)
            item.setBackground(color)
            item.setForeground(font)
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            if i == 1:
                item.setText(col1[list[1]])
            if i == 2:
                item.setText(col2[list[2]])
            if i == 3 and list[1] == 2:
                item.setText(f'{list[3]}s')
            if (i == 3 and list[1] == 0) or (i == 3 and list[1] == 3):
                item.setText(f'{list[3]}')
        QApplication.processEvents()

    def DIDO_itemOperation(self, list):
        '''
        :param list = [row,state,result,operationTime]
        :param row: 进行操作的单元行
        :param state: 测试状态 ['无需检测','正在检测','检测完成','等待检测']
        :param result: 测试结果 col2 = ['','通过','未通过']
        :param operationTime: 测试时间
        :return:
        '''
        mTable = self.tableWidget_DIDO
        col1 = ['无需检测', '正在检测', '检测完成', '等待检测']
        col2 = ['', '通过', '未通过']
        font = QtGui.QColor(0, 0, 0)
        if list[1] == 0:
            color = QtGui.QColor(255, 255, 255)
            font = QtGui.QColor(197, 197, 197)
        elif list[1] == 1:
            color = QtGui.QColor(255, 255, 0)
        elif list[1] == 2 and list[2] == 1:
            color = QtGui.QColor(0, 255, 0)
        elif list[1] == 2 and list[2] == 2:
            color = QtGui.QColor(255, 0, 0)
            font = QtGui.QColor(255, 255, 255)
        elif list[1] == 3:
            color = QtGui.QColor(197, 197, 197)

        for i in range(4):
            item = mTable.item(list[0], i)
            item.setBackground(color)
            item.setForeground(font)
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            if i == 1:
                item.setText(col1[list[1]])
            if i == 2:
                item.setText(col2[list[2]])
            if i == 3 and list[1] == 2:
                item.setText(f'{list[3]}s')
            if (i == 3 and list[1] == 0) or (i == 3 and list[1] == 3):
                item.setText(f'{list[3]}')
        QApplication.processEvents()

    def CPU_itemOperation(self,list):
        '''
        :param list = [row,state,result,operationTime]
        :param row: 进行操作的单元行
        :param state: 测试状态 ['无需检测','正在检测','检测完成','等待检测']
        :param result: 测试结果 col2 = ['','通过','未通过']
        :param operationTime: 测试时间
        :return:
        '''
        mTable = self.tableWidget_CPU
        col1 = ['无需检测','正在检测','检测完成','等待检测']
        col2 = ['','通过','未通过']
        font = QtGui.QColor(0, 0, 0)
        if list[1] == 0:
            color = QtGui.QColor(255, 255, 255)
            font = QtGui.QColor(197, 197, 197)
        elif list[1] == 1:
            color = QtGui.QColor(255,255,0)
        elif list[1] == 2 and list[2] == 1:
            color = QtGui.QColor(0, 255, 0)
        elif list[1] == 2 and list[2] == 2:
            color = QtGui.QColor(255, 0, 0)
            font = QtGui.QColor(255, 255, 255)
        elif list[1] == 3:
            color = QtGui.QColor(197, 197, 197)

        for i in range(4):
            item = mTable.item(list[0], i)
            item.setBackground(color)
            item.setForeground(font)
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            if i == 1:
                item.setText(col1[list[1]])
            if i == 2:
                item.setText(col2[list[2]])
            if i == 3 and list[1] == 2:
                item.setText(f'{list[3]}s')
            if (i == 3 and list[1] == 0) or (i == 3 and list[1] == 3):
                item.setText(f'{list[3]}')
        QApplication.processEvents()

    def itemOperation(self,mTable,row,state,result,operationTime):
        '''

        :param mTable: 进行操作的表格
        :param row: 进行操作的单元行
        :param state: 测试状态 ['无需检测','正在检测','检测完成','等待检测']
        :param result: 测试结果 col2 = ['','通过','未通过']
        :param operationTime: 测试时间
        :return:
        '''
        col1 = ['无需检测','正在检测','检测完成','等待检测']
        col2 = ['','通过','未通过']
        font = QtGui.QColor(0, 0, 0)
        if state == 0:
            color = QtGui.QColor(255, 255, 255)
            font = QtGui.QColor(197, 197, 197)
        elif state == 1:
            color = QtGui.QColor(255,255,0)
        elif state == 2 and result == 1:
            color = QtGui.QColor(0, 255, 0)
        elif state == 2 and result == 2:
            color = QtGui.QColor(255, 0, 0)
            font = QtGui.QColor(255, 255, 255)
        elif state == 3:
            color = QtGui.QColor(197, 197, 197)

        for i in range(4):
            item = mTable.item(row, i)
            item.setBackground(color)
            item.setForeground(font)
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            if i == 1:
                item.setText(col1[state])
            if i == 2:
                item.setText(col2[result])
            if i == 3 and state == 2:
                item.setText(f'{operationTime}s')
            if (i == 3 and state == 0) or (i == 3 and state == 3):
                item.setText(f'{operationTime}')
        QApplication.processEvents()

    def option_pushButton13(self):
        self.lineEdit_SN.setText('S1223-001083')
        self.lineEdit_REV.setText('06')

    # def generateExcel(self, station, module):
    # def generateExcel(self, list):
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
    #     if (self.isAITestVol and not self.isAITestCur) or (not self.isAITestVol and self.isAITestCur):
    #         sheet.write_merge(self.AI_row, self.AI_row + 1 + self.AI_Channels, 0, 0, 'AI信号', leftTitle_style)
    #         self.AO_row = self.AI_row + 2 + self.AI_Channels
    #     elif self.isAITestVol and self.isAITestCur:
    #         sheet.write_merge(self.AI_row, self.AI_row + 1 + 2 * self.AI_Channels, 0, 0, 'AI信号', leftTitle_style)
    #         self.AO_row = self.AI_row + 2 + 2 * self.AI_Channels
    #     elif not self.isAITestVol and not self.isAITestCur:
    #         sheet.write_merge(self.AI_row, self.AI_row + 1, 0, 0, 'AI信号', leftTitle_style)
    #         self.AO_row = self.AI_row + 2
    #
    #
    #     if (self.isAOTestVol and not self.isAOTestCur) or (not self.isAOTestVol and self.isAOTestCur):
    #         sheet.write_merge(self.AO_row, self.AO_row + 1 + self.AO_Channels, 0, 0, 'AO信号', leftTitle_style)
    #         self.result_row = self.AO_row + 2 + self.AO_Channels
    #     elif self.isAOTestVol and self.isAOTestCur:
    #         sheet.write_merge(self.AO_row, self.AO_row + 1 + 2 * self.AO_Channels, 0, 0, 'AO信号', leftTitle_style)
    #         self.result_row = self.AO_row + 2 + 2 * self.AO_Channels
    #     elif not self.isAOTestVol and not self.isAOTestCur:
    #         sheet.write_merge(self.AO_row, self.AO_row + 1, 0, 0, 'AO信号', leftTitle_style)
    #         self.result_row = self.AO_row + 2
    #
    #     # sheet.write_merge(self.AO_row, self.AO_row + 1 + self.AO_Channels, 0, 0, 'AO信号', leftTitle_style)
    #     # self.result_row = self.AO_row + 2
    #     sheet.write_merge(self.result_row, self.result_row + 1, 0, 3, '整体检测结果', leftTitle_style)
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
    #     sheet.write(self.CPU_row  + 1, 3, '---', contentTitle_style)
    #     sheet.write_merge(self.CPU_row + 1, self.CPU_row + 1, 4, 5, 'U盘', contentTitle_style)
    #     sheet.write(self.CPU_row  + 1, 6, '---', contentTitle_style)
    #     sheet.write_merge(self.CPU_row + 1, self.CPU_row + 1, 7, 8, 'type-C', contentTitle_style)
    #     sheet.write(self.CPU_row  + 1, 9, '---', contentTitle_style)
    #     sheet.write_merge(self.CPU_row + 1, self.CPU_row + 1, 10, 11, 'RS232通讯', contentTitle_style)
    #     sheet.write(self.CPU_row  + 1, 12, '---', contentTitle_style)
    #     sheet.write_merge(self.CPU_row + 1, self.CPU_row + 1, 13, 14, 'RS485通讯', contentTitle_style)
    #     sheet.write(self.CPU_row  + 1, 15, '---', contentTitle_style)
    #     sheet.write_merge(self.CPU_row + 1, self.CPU_row + 1, 16, 17, 'CAN通讯(预留)', contentTitle_style)
    #     sheet.write(self.CPU_row  + 1, 18, '---', contentTitle_style)
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
    #                       'AI/AO信号检验要记录数据，电压和电流的精度为2‰以下为合格、电阻的精度0.5℃以下合格；其他测试项合格打“√”，否则打“×”',
    #                       contentTitle_style)
    #
    #     # 检测信息
    #     sheet.write_merge(self.result_row + 3, self.result_row + 3, 0, 1, '检验员：', contentTitle_style)
    #     sheet.write_merge(self.result_row + 3, self.result_row + 3, 2, 3, '555', contentTitle_style)
    #     sheet.write_merge(self.result_row + 3, self.result_row + 3, 4, 5, '检验日期：', contentTitle_style)
    #     sheet.write_merge(self.result_row + 3, self.result_row + 3, 6, 8, f'{time.strftime("%Y-%m-%d %H：%M：%S")}', contentTitle_style)
    #     sheet.write_merge(self.result_row + 3, self.result_row + 3, 9, 10, '审核：', contentTitle_style)
    #     sheet.write_merge(self.result_row + 3, self.result_row + 3, 11, 13, ' ', contentTitle_style)
    #     sheet.write_merge(self.result_row + 3, self.result_row + 3, 14, 15, '审核日期：', contentTitle_style)
    #     sheet.write_merge(self.result_row + 3, self.result_row + 3, 16, 18, ' ', contentTitle_style)
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
    #     if list[1] == 'DI':
    #         self.fillInDIData(list[0], book, sheet)
    #     elif list[1] == 'DO':
    #         self.fillInDOData(list[0], book, sheet)
    #     elif list[1] == 'AI':
    #         # # print('打印AI检测结果')
    #         self.fillInAIData(list[0], book, sheet)
    #     elif list[1] == 'AO':
    #         # # print('打印AI检测结果')
    #         self.fillInAOData(list[0], book, sheet)
    #
    #
    # # @abstractmethod
    # # def fillInDIData(self):
    # #     raise NotImplementedError()
    # def fillInDIData(self,station, book, sheet):
    #     #通过单元格样式
    #     # 为样式创建字体
    #     pass_font = xlwt.Font()
    #     # 字体类型
    #     pass_font.name = '宋'
    #     # 字体颜色
    #     pass_font.colour_index = 0
    #     # 字体大小，11为字号，20为衡量单位
    #     pass_font.height = 10 * 20
    #     # 字体加粗
    #     pass_font.bold = False
    #
    #     # 设置单元格对齐方式
    #     pass_alignment = xlwt.Alignment()
    #     # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    #     pass_alignment.horz = 0x02
    #     # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    #     pass_alignment.vert = 0x01
    #     # 设置自动换行
    #     pass_alignment.wrap = 1
    #
    #     # 设置边框
    #     pass_borders = xlwt.Borders()
    #     # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    #     # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    #     pass_borders.left = 5
    #     pass_borders.right = 5
    #     pass_borders.top = 5
    #     pass_borders.bottom = 5
    #
    #     # # 设置背景颜色
    #     # pattern = xlwt.Pattern()
    #     # # 设置背景颜色的模式
    #     # pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    #     # # 背景颜色
    #     # pattern.pattern_fore_colour = i
    #
    #     # # 初始化样式
    #     pass_style = xlwt.XFStyle()
    #     pass_style.borders = pass_borders
    #     pass_style.alignment = pass_alignment
    #     pass_style.font = pass_font
    #
    #
    #     #未通过单元格样式
    #     # 为样式创建字体
    #     fail_font = xlwt.Font()
    #     # 字体类型
    #     fail_font.name = '宋'
    #     # 字体颜色
    #     fail_font.colour_index = 1
    #     # 字体大小，11为字号，20为衡量单位
    #     fail_font.height = 10 * 20
    #     # 字体加粗
    #     fail_font.bold = False
    #
    #     # 设置单元格对齐方式
    #     fail_alignment = xlwt.Alignment()
    #     # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    #     fail_alignment.horz = 0x02
    #     # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    #     fail_alignment.vert = 0x01
    #     # 设置自动换行
    #     fail_alignment.wrap = 1
    #
    #     # 设置边框
    #     fail_borders = xlwt.Borders()
    #     # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    #     # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    #     fail_borders.left = 5
    #     fail_borders.right = 5
    #     fail_borders.top = 5
    #     fail_borders.bottom = 5
    #
    #     # 设置背景颜色
    #     fail_pattern = xlwt.Pattern()
    #     # 设置背景颜色的模式
    #     fail_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    #     # 背景颜色
    #     fail_pattern.pattern_fore_colour = 18
    #
    #     # # 初始化样式
    #     fail_style = xlwt.XFStyle()
    #     fail_style.borders = fail_borders
    #     fail_style.alignment = fail_alignment
    #     fail_style.font = fail_font
    #     fail_style.pattern = fail_pattern
    #
    #     # 提示警告单元格样式
    #     # 为样式创建字体
    #     warning_font = xlwt.Font()
    #     # 字体类型
    #     warning_font.name = '宋'
    #     # 字体颜色
    #     warning_font.colour_index = 2
    #     # 字体大小，11为字号，20为衡量单位
    #     warning_font.height = 12 * 20
    #     # 字体加粗
    #     warning_font.bold = True
    #
    #     # 设置单元格对齐方式
    #     warning_alignment = xlwt.Alignment()
    #     # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    #     warning_alignment.horz = 0x02
    #     # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    #     warning_alignment.vert = 0x01
    #     # 设置自动换行
    #     warning_alignment.wrap = 1
    #
    #     # 设置边框
    #     warning_borders = xlwt.Borders()
    #     # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    #     # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    #     warning_borders.left = 5
    #     warning_borders.right = 5
    #     warning_borders.top = 5
    #     warning_borders.bottom = 5
    #
    #     # 设置背景颜色
    #     pattern = xlwt.Pattern()
    #     # 设置背景颜色的模式
    #     pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    #     # 背景颜色
    #     pattern.pattern_fore_colour = 3
    #
    #     # # 初始化样式
    #     warning_style = xlwt.XFStyle()
    #     warning_style.borders = warning_borders
    #     warning_style.alignment = warning_alignment
    #     warning_style.font = warning_font
    #
    #     if station and self.testNum == 0:
    #         name_save = '合格'
    #     if station and self.testNum != 0:
    #         name_save = '部分合格'
    #     elif not station:
    #         name_save = '不合格'
    #
    #     if self.appearance:
    #         sheet.write(self.generalTest_row, 3, '√', pass_style)
    #     elif not self.appearance:
    #         sheet.write(self.generalTest_row, 3, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})外观存在瑕疵 '
    #
    #     if self.CAN_runLED and self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 12, '√', pass_style)
    #     elif not self.CAN_runLED and self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 12, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})CAN_RUN指示灯未亮 '
    #     if self.CAN_errorLED and self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 15, '√', pass_style)
    #     elif not self.CAN_errorLED and self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 15, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})CAN_ERROR指示灯未亮 '
    #     if self.runLED and self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 6, '√', pass_style)
    #     elif not self.runLED and self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 6, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})RUN指示灯未亮 '
    #     if self.errorLED and self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 9, '√', pass_style)
    #     elif not self.errorLED and self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 9, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})ERROE指示灯未亮 '
    #     # #填写信号类型、通道号、测试点数据
    #     #     if self.isAITestVol:
    #     #         all_row = 9 + 4 + 4 + (2 + self.AI_Channels) + 2  # CPU + DI + DO + AI + AO
    #     #         sheet.write_merge(self.AI_row + 2,self.AI_row + 1 + self.AI_Channels,1,1,'电压',pass_style)
    #     #         for i in range(self.AI_Channels):
    #     #             #通道号
    #     #             sheet.write(self.AI_row + 2 + i, 2, f'CH{i+1}', pass_style)
    #     #             for i in range(5):
    #     #                 #理论值
    #     #                 sheet.write(self.AI_row + 2 + i, 3 + 3 * i, f'{self.voltageTheory[i]}', pass_style)
    #     #                 #测试值
    #     #                 sheet.write(self.AI_row + 2 + i, 4 + 3 * i, f'{self.volReceValue[i]}', pass_style)
    #     #                 # 精度
    #     #                 if abs(self.volPrecision[i]) < 2:
    #     #                     sheet.write(self.AI_row + 2 + i, 5 + 3 * i, f'{self.volPrecision[i]}‰', pass_style)
    #     #                 else:
    #     #                     sheet.write(self.AI_row + 2 + i, 5 + 3 * i, f'{self.volPrecision[i]}‰', fail_style)
    #     #     if self.isAITestVol and self.isAITestCur:
    #     #         all_row = 9 + 4 + 4 + (2 + 2 * self.AI_Channels) + 2  # CPU + DI + DO + AI + AO
    #     #         sheet.write_merge(self.AI_row + 2 + self.AI_Channels,self.AI_row + 1 + 2 * self.AI_Channels,1,1,'电流',pass_style)
    #     #         for i in range(self.AI_Channels):
    #     #             #通道号
    #     #             sheet.write(self.AI_row + 6 + i, 2, f'CH{i+1}', pass_style)
    #     #             for i in range(5):
    #     #                 #理论值
    #     #                 sheet.write(self.AI_row + 6 + i, 3 + 3 * i, f'{self.currentTheory[i]}', pass_style)
    #     #                 #测试值
    #     #                 sheet.write(self.AI_row + 6 + i, 4 + 3 * i, f'{self.curReceValue[i]}', pass_style)
    #     #                 # 精度
    #     #                 if abs(self.curPrecision[i]) < 2:
    #     #                     sheet.write(self.AI_row + 6 + i, 5 + 3 * i, f'{self.curPrecision[i]}‰', pass_style)
    #     #                 else:
    #     #                     sheet.write(self.AI_row + 6 + i, 5 + 3 * i, f'{self.curPrecision[i]}‰', fail_style)
    #     #     if not self.isAITestVol and self.isAITestCur:
    #     #         all_row = 9 + 4 + 4 + (2 + self.AI_Channels) + 2  # CPU + DI + DO + AI + AO
    #     #         sheet.write_merge(self.AI_row + 2, self.AI_row + 1 + self.AI_Channels, 1, 1, '电流', pass_style)
    #     #         for i in range(self.AI_Channels):
    #     #             # 通道号
    #     #             sheet.write(self.AI_row + 2 + i, 2, f'CH{i + 1}', pass_style)
    #     #             for i in range(5):
    #     #                 # 理论值
    #     #                 sheet.write(self.AI_row + 2 + i, 3 + 3 * i, f'{self.currentTheory[i]}', pass_style)
    #     #                 # 测试值
    #     #                 sheet.write(self.AI_row + 2 + i, 4 + 3 * i, f'{self.curReceValue[i]}', pass_style)
    #     #                 # 精度
    #     #                 if abs(self.curPrecision[i]) < 2:
    #     #                     sheet.write(self.AI_row + 2 + i, 5 + 3 * i, f'{self.curPrecision[i]}‰', pass_style)
    #     #                 else:
    #     #                     sheet.write(self.AI_row + 2 + i, 5 + 3 * i, f'{self.curPrecision[i]}‰', fail_style)
    #     #     if not self.isAITestVol and not self.isAITestCur:
    #     #         all_row = 9 + 4 + 4 + 2 + 2  # CPU + DI + DO + AI + AO
    #     #填写通道状态
    #     # self.DODataCheck[0] = True
    #     for i in range(self.m_Channels):
    #         if i < 16:
    #             if self.DIDataCheck[i]:
    #                 sheet.write(self.DI_row + 1, 3 + i, '√', pass_style)
    #             elif not self.DIDataCheck[i]:
    #                 sheet.write(self.DI_row + 1, 3 + i, '×', fail_style)
    #                 self.errorNum += 1
    #                 self.errorInf += f'{self.errorNum})通道{i+1} 灯未亮 '
    #         else:
    #             if self.DIDataCheck[i]:
    #                 sheet.write(self.DI_row + 3, i - 13, '√', pass_style)
    #             elif not self.DIDataCheck[i]:
    #                 sheet.write(self.DI_row + 3, i - 13, '×', fail_style)
    #                 self.errorNum += 1
    #                 self.errorInf += f'{self.errorNum})通道{i + 1} 灯未亮 '
    #
    #     self.isDIPassTest = (((((self.isDIPassTest & self.isLEDRunOK) & self.isLEDErrOK) &
    #                           self.CAN_runLED) & self.CAN_errorLED) & self.appearance)
    #     # self.showInf(f'self.isLEDRunOK:{self.isLEDRunOK}')
    #     all_row = 9 + 4 + 4 + 2 + 2  # CPU + DI + DO + AI + AO
    #     if self.isDIPassTest and self.testNum == 0:
    #         name_save = '合格'
    #         sheet.write(self.generalTest_row + all_row + 1, 4, '■ 合格', pass_style)
    #         self.label.setStyleSheet(self.testState_qss['pass'])
    #         self.label.setText('通过')
    #     if self.isDIPassTest and self.testNum > 0:
    #         name_save = '部分合格'
    #         sheet.write(self.generalTest_row + all_row + 1, 4, '■ 部分合格', pass_style)
    #         sheet.write(self.generalTest_row + all_row + 1, 6, '------------------ 注意：有部分项目未测试！！！ ------------------', warning_style)
    #         self.label.setStyleSheet(self.testState_qss['testing'])
    #         self.label.setText('部分通过')
    #     elif not self.isDIPassTest:
    #         name_save = '不合格'
    #         sheet.write(self.generalTest_row + all_row + 2, 4, '■ 不合格', fail_style)
    #         sheet.write(self.generalTest_row + all_row + 2, 6, f'不合格原因：{self.errorInf}', warning_style)
    #         self.label.setStyleSheet(self.testState_qss['fail'])
    #         self.label.setText('未通过')
    #     book.save(self.saveDir + f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}.xls')

    # @abstractmethod
    # def fillInDOData(self, station, book, sheet):
    #     # 通过单元格样式
    #     # 为样式创建字体
    #     pass_font = xlwt.Font()
    #     # 字体类型
    #     pass_font.name = '宋'
    #     # 字体颜色
    #     pass_font.colour_index = 0
    #     # 字体大小，11为字号，20为衡量单位
    #     pass_font.height = 10 * 20
    #     # 字体加粗
    #     pass_font.bold = False
    #
    #     # 设置单元格对齐方式
    #     pass_alignment = xlwt.Alignment()
    #     # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    #     pass_alignment.horz = 0x02
    #     # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    #     pass_alignment.vert = 0x01
    #     # 设置自动换行
    #     pass_alignment.wrap = 1
    #
    #     # 设置边框
    #     pass_borders = xlwt.Borders()
    #     # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    #     # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    #     pass_borders.left = 5
    #     pass_borders.right = 5
    #     pass_borders.top = 5
    #     pass_borders.bottom = 5
    #
    #     # # 设置背景颜色
    #     # pattern = xlwt.Pattern()
    #     # # 设置背景颜色的模式
    #     # pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    #     # # 背景颜色
    #     # pattern.pattern_fore_colour = i
    #
    #     # # 初始化样式
    #     pass_style = xlwt.XFStyle()
    #     pass_style.borders = pass_borders
    #     pass_style.alignment = pass_alignment
    #     pass_style.font = pass_font
    #
    #     # 未通过单元格样式
    #     # 为样式创建字体
    #     fail_font = xlwt.Font()
    #     # 字体类型
    #     fail_font.name = '宋'
    #     # 字体颜色
    #     fail_font.colour_index = 1
    #     # 字体大小，11为字号，20为衡量单位
    #     fail_font.height = 10 * 20
    #     # 字体加粗
    #     fail_font.bold = False
    #
    #     # 设置单元格对齐方式
    #     fail_alignment = xlwt.Alignment()
    #     # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    #     fail_alignment.horz = 0x02
    #     # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    #     fail_alignment.vert = 0x01
    #     # 设置自动换行
    #     fail_alignment.wrap = 1
    #
    #     # 设置边框
    #     fail_borders = xlwt.Borders()
    #     # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    #     # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    #     fail_borders.left = 5
    #     fail_borders.right = 5
    #     fail_borders.top = 5
    #     fail_borders.bottom = 5
    #
    #     # 设置背景颜色
    #     fail_pattern = xlwt.Pattern()
    #     # 设置背景颜色的模式
    #     fail_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    #     # 背景颜色
    #     fail_pattern.pattern_fore_colour = 18
    #
    #     # # 初始化样式
    #     fail_style = xlwt.XFStyle()
    #     fail_style.borders = fail_borders
    #     fail_style.alignment = fail_alignment
    #     fail_style.font = fail_font
    #     fail_style.pattern = fail_pattern
    #
    #     # 提示警告单元格样式
    #     # 为样式创建字体
    #     warning_font = xlwt.Font()
    #     # 字体类型
    #     warning_font.name = '宋'
    #     # 字体颜色
    #     warning_font.colour_index = 2
    #     # 字体大小，11为字号，20为衡量单位
    #     warning_font.height = 12 * 20
    #     # 字体加粗
    #     warning_font.bold = True
    #
    #     # 设置单元格对齐方式
    #     warning_alignment = xlwt.Alignment()
    #     # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    #     warning_alignment.horz = 0x02
    #     # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    #     warning_alignment.vert = 0x01
    #     # 设置自动换行
    #     warning_alignment.wrap = 1
    #
    #     # 设置边框
    #     warning_borders = xlwt.Borders()
    #     # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    #     # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    #     warning_borders.left = 5
    #     warning_borders.right = 5
    #     warning_borders.top = 5
    #     warning_borders.bottom = 5
    #
    #     # 设置背景颜色
    #     pattern = xlwt.Pattern()
    #     # 设置背景颜色的模式
    #     pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    #     # 背景颜色
    #     pattern.pattern_fore_colour = 3
    #
    #     # # 初始化样式
    #     warning_style = xlwt.XFStyle()
    #     warning_style.borders = warning_borders
    #     warning_style.alignment = warning_alignment
    #     warning_style.font = warning_font
    #
    #     # if station and self.testNum == 0:
    #     #     name_save = '合格'
    #     # if station and self.testNum != 0:
    #     #     name_save = '部分合格'
    #     # elif not station:
    #     #     name_save = '不合格'
    #
    #     if self.appearance:
    #         sheet.write(self.generalTest_row, 3, '√', pass_style)
    #     elif not self.appearance:
    #         sheet.write(self.generalTest_row, 3, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})外观存在瑕疵 '
    #
    #     if self.CAN_runLED and self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 12, '√', pass_style)
    #     elif not self.CAN_runLED and self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 12, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})CAN_RUN指示灯未亮 '
    #     if self.CAN_errorLED and self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 15, '√', pass_style)
    #     elif not self.CAN_errorLED and self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 15, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})CAN_ERROR指示灯未亮 '
    #     if self.runLED and self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 6, '√', pass_style)
    #     elif not self.runLED and self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 6, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})RUN指示灯未亮 '
    #     if self.errorLED and self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 9, '√', pass_style)
    #     elif not self.errorLED and self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 9, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})ERROE指示灯未亮 '
    #
    #     # 填写通道状态
    #     # self.DODataCheck[0] = True
    #     for i in range(self.m_Channels):
    #         if i < 16:
    #             if self.DODataCheck[i]:
    #                 sheet.write(self.DO_row + 1, 3 + i, '√', pass_style)
    #             elif not self.DODataCheck[i]:
    #                 sheet.write(self.DO_row + 1, 3 + i, '×', fail_style)
    #                 self.errorNum += 1
    #                 self.errorInf += f'{self.errorNum}.通道{i + 1}灯未亮 '
    #         else:
    #             if self.DODataCheck[i]:
    #                 sheet.write(self.DO_row + 3, i - 13, '√', pass_style)
    #             elif not self.DODataCheck[i]:
    #                 sheet.write(self.DO_row + 3, i - 13, '×', fail_style)
    #                 self.errorNum += 1
    #                 self.errorInf += f'{self.errorNum}.通道{i + 1}灯未亮 '
    #
    #     self.isDOPassTest = (((((self.isDOPassTest & self.isLEDRunOK) & self.isLEDErrOK) &
    #                            self.CAN_runLED) & self.CAN_errorLED) & self.appearance)
    #     # self.showInf(f'self.isLEDRunOK:{self.isLEDRunOK}')
    #     all_row = 9 + 4 + 4 + 2 + 2
    #     if self.isDOPassTest and self.testNum == 0:
    #         name_save = '合格'
    #         sheet.write(self.generalTest_row + all_row + 1, 4, '■ 合格', pass_style)
    #         self.label.setStyleSheet(self.testState_qss['pass'])
    #         self.label.setText('通过')
    #     if self.isDOPassTest and self.testNum > 0:
    #         name_save = '部分合格'
    #         sheet.write(self.generalTest_row + all_row + 1, 4, '■ 部分合格', pass_style)
    #         sheet.write(self.generalTest_row + all_row + 1, 6,
    #                     '------------------ 注意：有部分项目未测试！！！ ------------------', warning_style)
    #         self.label.setStyleSheet(self.testState_qss['testing'])
    #         self.label.setText('部分通过')
    #     elif not self.isDOPassTest:
    #         name_save = '不合格'
    #         sheet.write(self.generalTest_row + all_row + 2, 4, '■ 不合格', fail_style)
    #         sheet.write(self.generalTest_row + all_row + 2, 6, f'不合格原因：{self.errorInf}', fail_style)
    #         self.label.setStyleSheet(self.testState_qss['fail'])
    #         self.label.setText('未通过')
    #
    #     book.save(self.saveDir + f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}.xls')
    #
    # # @abstractmethod
    # def fillInAIData(self,station, book, sheet):
    #     # 通过单元格样式
    #     # 为样式创建字体
    #     pass_font = xlwt.Font()
    #     # 字体类型
    #     pass_font.name = '宋'
    #     # 字体颜色
    #     pass_font.colour_index = 0
    #     # 字体大小，11为字号，20为衡量单位
    #     pass_font.height = 10 * 20
    #     # 字体加粗
    #     pass_font.bold = False
    #
    #     # 设置单元格对齐方式
    #     pass_alignment = xlwt.Alignment()
    #     # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    #     pass_alignment.horz = 0x02
    #     # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    #     pass_alignment.vert = 0x01
    #     # 设置自动换行
    #     pass_alignment.wrap = 1
    #
    #     # 设置边框
    #     pass_borders = xlwt.Borders()
    #     # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    #     # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    #     pass_borders.left = 5
    #     pass_borders.right = 5
    #     pass_borders.top = 5
    #     pass_borders.bottom = 5
    #
    #     # 设置背景颜色
    #     pass_pattern = xlwt.Pattern()
    #     # 设置背景颜色的模式
    #     pass_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    #     # 背景颜色
    #     pass_pattern.pattern_fore_colour = 3
    #
    #     # # 初始化样式
    #     pass_style = xlwt.XFStyle()
    #     pass_style.borders = pass_borders
    #     pass_style.alignment = pass_alignment
    #     pass_style.font = pass_font
    #     pass_style.pattern = pass_pattern
    #
    #     # 未通过单元格样式
    #     # 为样式创建字体
    #     fail_font = xlwt.Font()
    #     # 字体类型
    #     fail_font.name = '宋'
    #     # 字体颜色
    #     fail_font.colour_index = 1
    #     # 字体大小，11为字号，20为衡量单位
    #     fail_font.height = 10 * 20
    #     # 字体加粗
    #     fail_font.bold = True
    #
    #     # 设置单元格对齐方式
    #     fail_alignment = xlwt.Alignment()
    #     # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    #     fail_alignment.horz = 0x02
    #     # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    #     fail_alignment.vert = 0x01
    #     # 设置自动换行
    #     fail_alignment.wrap = 1
    #
    #     # 设置边框
    #     fail_borders = xlwt.Borders()
    #     # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    #     # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    #     fail_borders.left = 5
    #     fail_borders.right = 5
    #     fail_borders.top = 5
    #     fail_borders.bottom = 5
    #
    #     # 设置背景颜色
    #     fail_pattern = xlwt.Pattern()
    #     # 设置背景颜色的模式
    #     fail_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    #     # 背景颜色
    #     fail_pattern.pattern_fore_colour = 2
    #
    #     # # 初始化样式
    #     fail_style = xlwt.XFStyle()
    #     fail_style.borders = fail_borders
    #     fail_style.alignment = fail_alignment
    #     fail_style.font = fail_font
    #     fail_style.pattern = fail_pattern
    #
    #     # 提示警告单元格样式
    #     # 为样式创建字体
    #     warning_font = xlwt.Font()
    #     # 字体类型
    #     warning_font.name = '宋'
    #     # 字体颜色
    #     warning_font.colour_index = 2
    #     # 字体大小，11为字号，20为衡量单位
    #     warning_font.height = 10 * 20
    #     # 字体加粗
    #     warning_font.bold = True
    #
    #     # 设置单元格对齐方式
    #     warning_alignment = xlwt.Alignment()
    #     # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    #     warning_alignment.horz = 0x02
    #     # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    #     warning_alignment.vert = 0x01
    #     # 设置自动换行
    #     warning_alignment.wrap = 1
    #
    #     # 设置边框
    #     warning_borders = xlwt.Borders()
    #     # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    #     # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    #     warning_borders.left = 5
    #     warning_borders.right = 5
    #     warning_borders.top = 5
    #     warning_borders.bottom = 5
    #
    #     # 设置背景颜色
    #     warning_pattern = xlwt.Pattern()
    #     # 设置背景颜色的模式
    #     warning_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    #     # 背景颜色
    #     warning_pattern.pattern_fore_colour = 5
    #
    #     # # 初始化样式
    #     warning_style = xlwt.XFStyle()
    #     warning_style.borders = warning_borders
    #     warning_style.alignment = warning_alignment
    #     warning_style.font = warning_font
    #     warning_style.pattern = warning_pattern
    #
    #     # if station and self.testNum == 0:
    #     #     name_save = '合格'
    #     # if station and self.testNum != 0:
    #     #     name_save = '部分合格'
    #     # elif not station:
    #     #     name_save = '不合格'
    #
    #     if self.appearance:
    #         sheet.write(self.generalTest_row, 3, '√', pass_style)
    #     elif not self.appearance:
    #         sheet.write(self.generalTest_row, 3, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})外观存在瑕疵 '
    #     if not self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 6, '未检测', warning_style)
    #         sheet.write(self.generalTest_row, 9, '未检测', warning_style)
    #     if not self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 12, '未检测', warning_style)
    #         sheet.write(self.generalTest_row, 15, '未检测', warning_style)
    #     if self.CAN_runLED and self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 12, '√', pass_style)
    #     elif not self.CAN_runLED and self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 12, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})CAN_RUN指示灯未亮 '
    #     if self.CAN_errorLED and self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 15, '√', pass_style)
    #     elif not self.CAN_errorLED and self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 15, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})CAN_ERROR指示灯未亮 '
    #     if self.runLED and self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 6, '√', pass_style)
    #     elif not self.runLED and self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 6, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})RUN指示灯未亮 '
    #     if self.errorLED and self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 9, '√', pass_style)
    #     elif not self.errorLED and self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 9, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})ERROE指示灯未亮 '
    #     # 填写信号类型、通道号、测试点数据
    #     if self.isAITestVol:
    #         all_row = 9 + 4 + 4 + (2 + self.AI_Channels) + 2  # CPU + DI + DO + AI + AO
    #         sheet.write_merge(self.AI_row + 2, self.AI_row + 1 + self.AI_Channels, 1, 1, '电压', pass_style)
    #         for i in range(5):
    #             for j in range(self.AI_Channels):
    #                 # 通道号
    #                 sheet.write_merge(self.AI_row + 2 + j, self.AI_row + 2 + j, 2, 3, f'CH{j + 1}', pass_style)
    #                 # 理论值
    #                 sheet.write(self.AI_row + 2 + j, 3 + 3 * i + 1, f'{self.voltageTheory[i]}', pass_style)
    #                 # 测试值
    #                 sheet.write(self.AI_row + 2 + j, 4 + 3 * i + 1, f'{self.volReceValue[i][j]}', pass_style)
    #                 # 精度
    #                 if abs(self.volPrecision[i][j]) < 2:
    #                     sheet.write(self.AI_row + 2 + j, 5 + 3 * i + 1, f'{self.volPrecision[i][j]}‰', pass_style)
    #                 else:
    #                     self.errorNum += 1
    #                     self.errorInf += f'{self.errorNum})测试点{i+1}AI通道{j+1}电压精度超出范围 '
    #                     sheet.write(self.AI_row + 2 + j, 5 + 3 * i + 1, f'{self.volPrecision[i][j]}‰', fail_style)
    #     if self.isAITestVol and self.isAITestCur:
    #         all_row = 9 + 4 + 4 + (2 + 2 * self.AI_Channels) + 2  # CPU + DI + DO + AI + AO
    #         sheet.write_merge(self.AI_row + 2 + self.AI_Channels, self.AI_row + 1 + 2 * self.AI_Channels, 1, 1,
    #                           '电流', pass_style)
    #         for i in range(5):
    #             for j in range(self.AI_Channels):
    #                 # 通道号
    #                 sheet.write_merge(self.AI_row + 2 +self.AI_Channels + j,self.AI_row + 6 + j, 2, 3, f'CH{j + 1}', pass_style)
    #                 # 理论值
    #                 sheet.write(self.AI_row + 2 +self.AI_Channels + j, 3 + 3 * i + 1, f'{self.currentTheory[i]}', pass_style)
    #                 # 测试值
    #                 sheet.write(self.AI_row + 2 +self.AI_Channels + j, 4 + 3 * i + 1, f'{self.curReceValue[i][j]}', pass_style)
    #                 # 精度
    #                 if abs(self.curPrecision[i][j]) < 2:
    #                     sheet.write(self.AI_row + 2 +self.AI_Channels + j, 5 + 3 * i + 1, f'{self.curPrecision[i][j]}‰', pass_style)
    #                 else:
    #                     self.errorNum += 1
    #                     self.errorInf += f'{self.errorNum})测试点{i+1}AI通道{j + 1}电流精度超出范围 '
    #                     sheet.write(self.AI_row + 2 +self.AI_Channels + j, 5 + 3 * i + 1, f'{self.curPrecision[i][j]}‰', fail_style)
    #     if not self.isAITestVol and self.isAITestCur:
    #         all_row = 9 + 4 + 4 + (2 + self.AI_Channels) + 2  # CPU + DI + DO + AI + AO
    #         sheet.write_merge(self.AI_row + 2, self.AI_row + 1 + self.AI_Channels, 1, 1, '电流', pass_style)
    #         for i in range(5):
    #             for j in range(self.AI_Channels):
    #                 # 通道号
    #                 sheet.write_merge(self.AI_row + 2 + j, self.AI_row + 2 + j, 2, 3, f'CH{j + 1}', pass_style)
    #                 # 理论值
    #                 sheet.write(self.AI_row + 2 + j, 3 + 3 * i + 1, f'{self.currentTheory[i]}', pass_style)
    #                 # 测试值
    #                 # # print(j)
    #                 sheet.write(self.AI_row + 2 + j, 4 + 3 * i + 1, f'{self.curReceValue[i][j]}', pass_style)
    #                 # 精度
    #                 if abs(self.curPrecision[i][j]) < 2:
    #                     sheet.write(self.AI_row + 2 + j, 5 + 3 * i + 1, f'{self.curPrecision[i][j]}‰', pass_style)
    #                 else:
    #                     self.errorNum += 1
    #                     self.errorInf += f'{self.errorNum})测试点{i+1}AI通道{j + 1}电流精度超出范围'
    #                     sheet.write(self.AI_row + 2 + j, 5 + 3 * i + 1, f'{self.curPrecision[i][j]}‰', fail_style)
    #
    #     if not self.isAITestVol and not self.isAITestCur:
    #         all_row = 9 + 4 + 4 + 2 + 2  # CPU + DI + DO + AI + AO
    #     # print(f'self.isAIPassTest:{self.isAIPassTest}')
    #     self.isAIPassTest = (((((self.isAIPassTest & self.isLEDRunOK) & self.isLEDErrOK) & self.CAN_runLED) & self.CAN_errorLED) & self.appearance)
    #     # self.showInf(f'self.isLEDRunOK:{self.isLEDRunOK}')
    #     # print(f'self.isAIPassTest:{self.isAIPassTest}')
    #     # print(f'self.isLEDRunOK:{self.isLEDRunOK}')
    #     # print(f'self.isLEDErrOK:{self.isLEDErrOK}')
    #     # print(f'self.CAN_runLED:{self.CAN_runLED}')
    #     # print(f'self.CAN_errorLED:{self.CAN_errorLED}')
    #     # print(f'self.appearance:{self.appearance}')
    #     # print(f'self.testNum:{self.testNum}')
    #     # self.showInf(f'self.testNum:{self.testNum}')
    #     name_save = ''
    #     if self.isAIPassTest and self.testNum == 0:
    #         name_save = '合格'
    #         sheet.write(self.generalTest_row + all_row + 1, 4, '■ 合格', pass_style)
    #         sheet.write(self.generalTest_row + all_row + 1, 6,
    #                     '------------------ 全部项目测试通过！！！ ------------------', pass_style)
    #         self.label.setStyleSheet(self.testState_qss['pass'])
    #         self.label.setText('全部通过')
    #     elif self.isAIPassTest and self.testNum > 0:
    #         name_save = '部分合格'
    #         sheet.write(self.generalTest_row + all_row + 1, 4, '■ 部分合格', pass_style)
    #         sheet.write(self.generalTest_row + all_row + 1, 6,
    #                     '------------------ 注意：有部分项目未测试！！！ ------------------', warning_style)
    #         self.label.setStyleSheet(self.testState_qss['testing'])
    #         self.label.setText('部分通过')
    #     elif not self.isAIPassTest:
    #         name_save = '不合格'
    #         sheet.write(self.generalTest_row + all_row + 2, 4, '■ 不合格', fail_style)
    #         sheet.write(self.generalTest_row + all_row + 2, 6, f'不合格原因：{self.errorInf}', fail_style)
    #         self.label.setStyleSheet(self.testState_qss['fail'])
    #         self.label.setText('未通过')
    #
    #     book.save(self.saveDir + f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}.xls')
    #
    # # @abstractmethod
    # def fillInAOData(self,station, book, sheet):
    #     # 通过单元格样式
    #     # 为样式创建字体
    #     pass_font = xlwt.Font()
    #     # 字体类型
    #     pass_font.name = '宋'
    #     # 字体颜色
    #     pass_font.colour_index = 0
    #     # 字体大小，11为字号，20为衡量单位
    #     pass_font.height = 10 * 20
    #     # 字体加粗
    #     pass_font.bold = False
    #
    #     # 设置单元格对齐方式
    #     pass_alignment = xlwt.Alignment()
    #     # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    #     pass_alignment.horz = 0x02
    #     # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    #     pass_alignment.vert = 0x01
    #     # 设置自动换行
    #     pass_alignment.wrap = 1
    #
    #     # 设置边框
    #     pass_borders = xlwt.Borders()
    #     # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    #     # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    #     pass_borders.left = 5
    #     pass_borders.right = 5
    #     pass_borders.top = 5
    #     pass_borders.bottom = 5
    #
    #     # 设置背景颜色
    #     pass_pattern = xlwt.Pattern()
    #     # 设置背景颜色的模式
    #     pass_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    #     # 背景颜色
    #     pass_pattern.pattern_fore_colour = 3
    #
    #     # # 初始化样式
    #     pass_style = xlwt.XFStyle()
    #     pass_style.borders = pass_borders
    #     pass_style.alignment = pass_alignment
    #     pass_style.font = pass_font
    #     pass_style.pattern = pass_pattern
    #
    #     # 未通过单元格样式
    #     # 为样式创建字体
    #     fail_font = xlwt.Font()
    #     # 字体类型
    #     fail_font.name = '宋'
    #     # 字体颜色
    #     fail_font.colour_index = 1
    #     # 字体大小，11为字号，20为衡量单位
    #     fail_font.height = 10 * 20
    #     # 字体加粗
    #     fail_font.bold = True
    #
    #     # 设置单元格对齐方式
    #     fail_alignment = xlwt.Alignment()
    #     # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    #     fail_alignment.horz = 0x02
    #     # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    #     fail_alignment.vert = 0x01
    #     # 设置自动换行
    #     fail_alignment.wrap = 1
    #
    #     # 设置边框
    #     fail_borders = xlwt.Borders()
    #     # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    #     # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    #     fail_borders.left = 5
    #     fail_borders.right = 5
    #     fail_borders.top = 5
    #     fail_borders.bottom = 5
    #
    #     # 设置背景颜色
    #     fail_pattern = xlwt.Pattern()
    #     # 设置背景颜色的模式
    #     fail_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    #     # 背景颜色
    #     fail_pattern.pattern_fore_colour = 2
    #
    #     # # 初始化样式
    #     fail_style = xlwt.XFStyle()
    #     fail_style.borders = fail_borders
    #     fail_style.alignment = fail_alignment
    #     fail_style.font = fail_font
    #     fail_style.pattern = fail_pattern
    #
    #     # 提示警告单元格样式
    #     # 为样式创建字体
    #     warning_font = xlwt.Font()
    #     # 字体类型
    #     warning_font.name = '宋'
    #     # 字体颜色
    #     warning_font.colour_index = 2
    #     # 字体大小，11为字号，20为衡量单位
    #     warning_font.height = 10 * 20
    #     # 字体加粗
    #     warning_font.bold = True
    #
    #     # 设置单元格对齐方式
    #     warning_alignment = xlwt.Alignment()
    #     # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    #     warning_alignment.horz = 0x02
    #     # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    #     warning_alignment.vert = 0x01
    #     # 设置自动换行
    #     warning_alignment.wrap = 1
    #
    #     # 设置边框
    #     warning_borders = xlwt.Borders()
    #     # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    #     # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    #     warning_borders.left = 5
    #     warning_borders.right = 5
    #     warning_borders.top = 5
    #     warning_borders.bottom = 5
    #
    #     # 设置背景颜色
    #     warning_pattern = xlwt.Pattern()
    #     # 设置背景颜色的模式
    #     warning_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    #     # 背景颜色
    #     warning_pattern.pattern_fore_colour = 5
    #
    #     # # 初始化样式
    #     warning_style = xlwt.XFStyle()
    #     warning_style.borders = warning_borders
    #     warning_style.alignment = warning_alignment
    #     warning_style.font = warning_font
    #     warning_style.pattern = warning_pattern
    #
    #     # if station and self.testNum == 0:
    #     #     name_save = '合格'
    #     # if station and self.testNum != 0:
    #     #     name_save = '部分合格'
    #     # elif not station:
    #     #     name_save = '不合格'
    #
    #     if self.appearance:
    #         sheet.write(self.generalTest_row, 3, '√', pass_style)
    #     elif not self.appearance:
    #         sheet.write(self.generalTest_row, 3, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})外观存在瑕疵 '
    #     if not self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 6, '未检测', warning_style)
    #         sheet.write(self.generalTest_row, 9, '未检测', warning_style)
    #     if not self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 12, '未检测', warning_style)
    #         sheet.write(self.generalTest_row, 15, '未检测', warning_style)
    #     if self.CAN_runLED and self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 12, '√', pass_style)
    #     elif not self.CAN_runLED and self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 12, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})CAN_RUN指示灯未亮 '
    #     if self.CAN_errorLED and self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 15, '√', pass_style)
    #     elif not self.CAN_errorLED and self.isTestCANRunErr:
    #         sheet.write(self.generalTest_row, 15, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})CAN_ERROR指示灯未亮 '
    #     if self.runLED and self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 6, '√', pass_style)
    #     elif not self.runLED and self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 6, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})RUN指示灯未亮 '
    #     if self.errorLED and self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 9, '√', pass_style)
    #     elif not self.errorLED and self.isTestRunErr:
    #         sheet.write(self.generalTest_row, 9, '×', fail_style)
    #         self.errorNum += 1
    #         self.errorInf += f'{self.errorNum})ERROE指示灯未亮 '
    #     # 填写信号类型、通道号、测试点数据
    #     if self.isAOTestVol:
    #         all_row = 9 + 4 + 4 + (2 + self.AO_Channels) + 2  # CPU + DI + DO + AO + AI
    #         sheet.write_merge(self.AO_row + 2, self.AO_row + 1 + self.AO_Channels, 1, 1, '电压', pass_style)
    #         for i in range(5):
    #             for j in range(self.AO_Channels):
    #                 # 通道号
    #                 sheet.write_merge(self.AO_row + 2 + j, self.AO_row + 2 + j, 2, 3, f'CH{j + 1}', pass_style)
    #                 # 理论值
    #                 sheet.write(self.AO_row + 2 + j, 3 + 3 * i + 1, f'{self.voltageTheory[i]}', pass_style)
    #                 # 测试值
    #                 sheet.write(self.AO_row + 2 + j, 4 + 3 * i + 1, f'{self.volReceValue[i][j]}', pass_style)
    #                 # 精度
    #                 if abs(self.volPrecision[i][j]) < 2:
    #                     sheet.write(self.AO_row + 2 + j, 5 + 3 * i + 1, f'{self.volPrecision[i][j]}‰', pass_style)
    #                 else:
    #                     self.errorNum += 1
    #                     self.errorInf += f'{self.errorNum})AO通道{i + 1}电压精度超出范围 '
    #                     sheet.write(self.AO_row + 2 + j, 5 + 3 * i + 1, f'{self.volPrecision[i][j]}‰', fail_style)
    #     if self.isAOTestVol and self.isAOTestCur:
    #         all_row = 9 + 4 + 4 + (2 + 2 * self.AO_Channels) + 2  # CPU + DI + DO + AO + AI
    #         sheet.write_merge(self.AO_row + 2 + self.AO_Channels, self.AO_row + 1 + 2 * self.AO_Channels, 1, 1,
    #                           '电流', pass_style)
    #         for i in range(5):
    #             for j in range(self.AO_Channels):
    #                 # 通道号
    #                 sheet.write_merge(self.AO_row + 2 + self.AO_Channels + j, self.AO_row + 6 + j, 2, 3, f'CH{j + 1}',
    #                                   pass_style)
    #                 # 理论值
    #                 sheet.write(self.AO_row + 2 + self.AO_Channels + j, 3 + 3 * i + 1, f'{self.currentTheory[i]}',
    #                             pass_style)
    #                 # 测试值
    #                 sheet.write(self.AO_row + 2 + self.AO_Channels + j, 4 + 3 * i + 1, f'{self.curReceValue[i][j]}',
    #                             pass_style)
    #                 # 精度
    #                 if abs(self.curPrecision[i][j]) < 2:
    #                     sheet.write(self.AO_row + 2 + self.AO_Channels + j, 5 + 3 * i + 1,
    #                                 f'{self.curPrecision[i][j]}‰', pass_style)
    #                 else:
    #                     self.errorNum += 1
    #                     self.errorInf += f'{self.errorNum})测试点{i+1}AO通道{j + 1}电流精度超出范围 '
    #                     sheet.write(self.AO_row + 2 + self.AO_Channels + j, 5 + 3 * i + 1,
    #                                 f'{self.curPrecision[i][j]}‰', fail_style)
    #     if not self.isAOTestVol and self.isAOTestCur:
    #         all_row = 9 + 4 + 4 + (2 + self.AO_Channels) + 2  # CPU + DI + DO + AO + AI
    #         sheet.write_merge(self.AO_row + 2, self.AO_row + 1 + self.AO_Channels, 1, 1, '电流', pass_style)
    #         for i in range(5):
    #             for j in range(self.AO_Channels):
    #                 # 通道号
    #                 sheet.write_merge(self.AO_row + 2 + j, self.AO_row + 2 + j, 2, 3, f'CH{j + 1}', pass_style)
    #                 # 理论值
    #                 sheet.write(self.AO_row + 2 + j, 3 + 3 * i + 1, f'{self.currentTheory[i]}', pass_style)
    #                 # 测试值
    #                 # # print(j)
    #                 sheet.write(self.AO_row + 2 + j, 4 + 3 * i + 1, f'{self.curReceValue[i][j]}', pass_style)
    #                 # 精度
    #                 if abs(self.curPrecision[i][j]) < 2:
    #                     sheet.write(self.AO_row + 2 + j, 5 + 3 * i + 1, f'{self.curPrecision[i][j]}‰', pass_style)
    #                 else:
    #                     self.errorNum += 1
    #                     self.errorInf += f'{self.errorNum})测试点{i+1}AO通道{j + 1}电流精度超出范围'
    #                     sheet.write(self.AO_row + 2 + j, 5 + 3 * i + 1, f'{self.curPrecision[i][j]}‰', fail_style)
    #
    #     if not self.isAOTestVol and not self.isAOTestCur:
    #         all_row = 9 + 4 + 4 + 2 + 2  # CPU + DI + DO + AI + AO
    #     # print(f'self.isAOPassTest:{self.isAOPassTest}')
    #     self.isAOPassTest = (((((self.isAOPassTest & self.isLEDRunOK) & self.isLEDErrOK) & self.CAN_runLED) & self.CAN_errorLED) & self.appearance)
    #     # self.showInf(f'self.isLEDRunOK:{self.isLEDRunOK}')
    #     # print(f'self.isAOPassTest:{self.isAOPassTest}')
    #     # print(f'self.isLEDRunOK:{self.isLEDRunOK}')
    #     # print(f'self.isLEDErrOK:{self.isLEDErrOK}')
    #     # print(f'self.CAN_runLED:{self.CAN_runLED}')
    #     # print(f'self.CAN_errorLED:{self.CAN_errorLED}')
    #     # print(f'self.appearance:{self.appearance}')
    #     # print(f'self.testNum:{self.testNum}')
    #     # self.showInf(f'self.testNum:{self.testNum}')
    #     name_save = ''
    #     if self.isAOPassTest and self.testNum == 0:
    #         name_save = '合格'
    #         sheet.write(self.generalTest_row + all_row + 1, 4, '■ 合格', pass_style)
    #         sheet.write(self.generalTest_row + all_row + 1, 6,
    #                     '------------------ 全部项目测试通过！！！ ------------------', pass_style)
    #         self.label.setStyleSheet(self.testState_qss['pass'])
    #         self.label.setText('全部通过')
    #     elif self.isAOPassTest and self.testNum > 0:
    #         name_save = '部分合格'
    #         sheet.write(self.generalTest_row + all_row + 1, 4, '■ 部分合格', pass_style)
    #         sheet.write(self.generalTest_row + all_row + 1, 6,
    #                     '------------------ 注意：有部分项目未测试！！！ ------------------', warning_style)
    #         self.label.setStyleSheet(self.testState_qss['testing'])
    #         self.label.setText('部分通过')
    #     elif not self.isAOPassTest:
    #         name_save = '不合格'
    #         sheet.write(self.generalTest_row + all_row + 2, 4, '■ 不合格', fail_style)
    #         sheet.write(self.generalTest_row + all_row + 2, 6, f'不合格原因：{self.errorInf}', fail_style)
    #         self.label.setStyleSheet(self.testState_qss['fail'])
    #         self.label.setText('未通过')
    #
    #     book.save(self.saveDir + f'/{name_save}{self.module_type}_{time.strftime("%Y%m%d%H%M%S")}.xls')

    def initPara(self):
        # 生成表格的相关标识
        # DI测试是否通过
        self.isDIPassTest = True
        # DO测试是否通过
        self.isDOPassTest = True
        # DO通道数据初始化
        self.DO_channelData = 0
        # DI通道数据初始化
        self.DI_channelData = 0
        # DO通道错误数据记录
        self.DODataCheck = [True for i in range(32)]
        # DI通道错误数据记录
        self.DIDataCheck = [True for i in range(32)]

        # 是否标定
        self.isCalibrate = False
        # 是否标定电压
        self.isCalibrateVol = False
        # 是否标定电流
        self.isCalibrateCur = False

        # 是否检测
        self.isCalibrate = False
        # 是否检测AI电压
        self.isAITestVol = False
        # 是否检测AI电流
        self.isAITestCur = False
        # 是否检测AO电压
        self.isAOTestVol = False
        # 是否检测AO电流
        self.isAOTestCur = False
        # 是否检测CAN_Run+CAN_Error
        self.isTestCANRunErr = False
        # 是否检测Run+Error
        self.isTestRunErr = False
        # AI测试是否通过
        self.isAIPassTest = True
        self.isAIVolPass = True
        self.isAICurPass = True
        # AO测试是否通过
        self.isAOPassTest = True
        self.isAOVolPass = True
        self.isAOCurPass = True

        self.isLEDRunOK = True
        self.isLEDErrOK = True
        self.isLEDPass = True

        # 测试总数
        self.testNum = 0
        # 测试类型
        self.testType = {}
        # 模块信息
        # inf = {'AO2': '待检AO模块', 'AI1': '配套AI模块', 'AI2': '待检AI模块', 'AO1': '配套AO模块',
        #        'DO2': '待检DO模块', 'DI1': '配套DI模块', 'DI2': '待检DI模块', 'DO1': '配套DO模块'}
        self.module_1 = ''
        self.module_2 = ''
        # 通道数
        self.m_Channels = 0
        self.AI_Channels = 0
        self.AO_Channels = 0
        # 发送的数据
        self.ubyte_array_transmit = c_ubyte * 8
        self.m_transmitData = self.ubyte_array_transmit(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
        # 接收的数据
        self.ubyte_array_receive = c_ubyte * 8
        self.m_receiveData = self.ubyte_array_receive(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
        self.CAN_errorLED = True
        self.CAN_runLED = True
        self.errorLED = True
        self.runLED = True
        self.errorNum = 0
        self.errorInf = ''
        self.isAllScreen = False
        self.volReceValue = [0, 0, 0, 0, 0]
        self.curReceValue = [0, 0, 0, 0, 0]
        self.volPrecision = [0, 0, 0, 0, 0]
        self.curPrecision = [0, 0, 0, 0, 0]
        self.chPrecision = [0, 0, 0, 0]

        self.subRunFlag = True


def CAN_init(CAN_channel:int):
    CAN_option.close(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX)
    time.sleep(0.1)
    QApplication.processEvents()

    if not CAN_option.connect(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX, CAN_channel):
        # self.showMessageBox(['CAN设备存在问题', 'CAN设备开启失败，请检查CAN设备！'])
        # self.CANFail()
        return [False,['CAN设备存在问题', f'CAN通道{CAN_channel}开启失败，请检查CAN设备！']]

    if not CAN_option.init(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX, CAN_channel, init_config):
        # self.showMessageBox(['CAN设备存在问题', 'CAN通道初始化失败，请检查CAN设备！'])
        # self.CANFail()
        return [False,['CAN设备存在问题', f'CAN通道{CAN_channel}初始化失败，请检查CAN设备！']]

    if not CAN_option.start(CAN_option.VCI_USB_CAN_2, CAN_option.DEV_INDEX, CAN_channel):
        # self.showMessageBox(['CAN设备存在问题','CAN通道打开失败，请检查CAN设备！'])
        # self.CANFail()
        return [False,['CAN设备存在问题',f'CAN通道{CAN_channel}打开失败，请检查CAN设备！']]
    return [True,['','']]


def mainThreadRunning():
    global isMainRunning
    if not isMainRunning:
        return False
    return True
def canInit_thread():
    q.put(CAN_init(1))
    event_canInit.set()

def strToASCII(str):
    ascii_list = []
    for char in str:
        ascii_code = int(ord(char))
        ascii_list.append(ascii_code)

    # print(ascii_list)
    return ascii_list

def ASCIIToHex(list):
    hex_list = [0x00 for x in range(len(list))]
    # print(f'hex_list = {hex_list}')
    num = 0
    for n in list:
        # print(type(hex_list[num]))
        hex_list[num] = hex(n)
        # print(type(hex_list[num]))
        num += 1
    # print(f'hex_list = {hex_list}')
    return hex_list

def write3codeToPLC(addr,code_list):
    ubyte_transmit = c_ubyte * 8
    transmit_inf = ubyte_transmit(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
    #写入PN
    code_pn = code_list[0]
    # print(f'code_pn = {code_pn}')
    num1 = int(len(code_pn) / 4)
    num2 = len(code_pn) % 4
    if num1 != 0:
        for i in range(num1):
            transmit_inf = ubyte_transmit(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
            transmit_inf = [0x23, 0xf8, 0x5f,  0x00+(i+1), 0x00+code_pn[i*4+0], 0x00+code_pn[i*4+1],
                            0x00+code_pn[i*4+2], 0x00+code_pn[i*4+3]]
            # transmit_inf = [0x23, 0xf7, 0x5f, 0x01, 0x53, 0x31, 0x32, 0x32]
            bool_pn, rece_pn = CAN_option.transmitCAN(0x602, transmit_inf,0)
            if not bool_pn:
                return False
    if num2 != 0:
        transmit_inf = ubyte_transmit(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
        if num2 == 1:
            transmit_inf[0] = 0x2f
            transmit_inf[4] = code_pn[num1*4+0]
        elif num2 == 2:
            transmit_inf[0] = 0x2b
            transmit_inf[4] = 0x00 + code_pn[num1 * 4 + 0]
            transmit_inf[5] = 0x00 + code_pn[num1 * 4 + 1]
        elif num2 == 3:
            transmit_inf[0] = 0x27
            transmit_inf[4] = 0x00 + code_pn[num1 * 4 + 0]
            transmit_inf[5] = 0x00 + code_pn[num1 * 4 + 1]
            transmit_inf[6] = 0x00 + code_pn[num1 * 4 + 2]

        transmit_inf[1] = 0xf8
        transmit_inf[2] = 0x5f
        transmit_inf[3] = num1+1
        bool_pn,rece_pn = CAN_option.transmitCAN(0x602,transmit_inf,0)
        if not bool_pn:
            return False

    # 写入SN
    code_sn = code_list[1]
    # print('code_sn = ',code_sn)
    num1 = int(len(code_sn) / 4)
    num2 = len(code_sn) % 4
    if num1 != 0:
        for i in range(num1):
            transmit_inf = ubyte_transmit(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
            transmit_inf = [0x23, 0xf7, 0x5f, 0x00+(i+1), 0x00 + code_sn[i * 4 + 0], 0x00+code_sn[i * 4 + 1],
                            0x00 + code_sn[i * 4 + 2], 0x00 + code_sn[i * 4 + 3]]
            # transmit_inf = [0x23, 0xf7, 0x5f, 0x01, 0x53, 0x31,0x32,0x32]
            bool_sn, rece_sn = CAN_option.transmitCAN(0x602, transmit_inf,0)
            if not bool_sn:
                return False
    if num2 != 0:
        transmit_inf = ubyte_transmit(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
        if num2 == 1:
            transmit_inf[0] = 0x2f
            transmit_inf[4] = 0x00 + code_sn[num1 * 4 + 0]
        elif num2 == 2:
            transmit_inf[0] = 0x2b
            transmit_inf[4] = 0x00 + code_sn[num1 * 4 + 0]
            transmit_inf[5] = 0x00 + code_sn[num1 * 4 + 1]
        elif num2 == 3:
            transmit_inf[0] = 0x27
            transmit_inf[4] = 0x00 + code_sn[num1 * 4 + 0]
            transmit_inf[5] = 0x00 + code_sn[num1 * 4 + 1]
            transmit_inf[6] = 0x00 + code_sn[num1 * 4 + 2]

        transmit_inf[1] = 0xf7
        transmit_inf[2] = 0x5f
        transmit_inf[3] = 0x00 + (num1 + 1)
        bool_sn, rece_sn = CAN_option.transmitCAN(0x602, transmit_inf,0)
        if not bool_sn:
            return False
    # print('code_sn = ', code_sn)
    # 写入REV
    code_rev = code_list[2]
    num1 = int(len(code_rev) / 4)
    num2 = len(code_rev) % 4
    if num1 != 0:
        for i in range(num1):

            transmit_inf = ubyte_transmit(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
            transmit_inf = [0x23, 0xf9, 0x5f, 0x00 + (i+1), 0x00 + code_rev[i * 4 + 0], 0x00 + code_rev[i * 4 + 1],
                            0x00 + code_rev[i * 4 + 2], 0x00 + code_rev[i * 4 + 3]]
            bool_rev, rece_rev = CAN_option.transmitCAN(0x600 + addr, transmit_inf,0)
            if not bool_rev:
                return False
    if num2 != 0:
        transmit_inf = ubyte_transmit(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)
        if num2 == 1:
            transmit_inf[0] = 0x2f
            transmit_inf[4] = 0x00 + code_rev[num1 * 4 + 0]
        elif num2 == 2:
            transmit_inf[0] = 0x2b
            transmit_inf[4] = 0x00 + code_rev[num1 * 4 + 0]
            transmit_inf[5] = 0x00 + code_rev[num1 * 4 + 1]
        elif num2 == 3:
            transmit_inf[0] = 0x27
            transmit_inf[4] = 0x00 + code_rev[num1 * 4 + 0]
            transmit_inf[5] = 0x00 + code_rev[num1 * 4 + 1]
            transmit_inf[6] = 0x00 + code_rev[num1 * 4 + 2]

        transmit_inf[1] = 0xf9
        transmit_inf[2] = 0x5f
        transmit_inf[3] = 0x00 + (num1 + 1)
        bool_rev, rece_rev = CAN_option.transmitCAN(0x600 + addr, transmit_inf,0)
        if not bool_rev:
            return False

    return True

def get3codeFromPLC(addr):
    ubyte_transmit = c_ubyte * 8
    transmit_inf = ubyte_transmit(0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00)

    #读设备PN
    pn_list = []
    flag_pn = True
    for i in range(6):
        transmit_inf[0] = 0x40
        transmit_inf[1] = 0xf8
        transmit_inf[2] = 0x5f
        transmit_inf[3] = 0x00 + i + 1
        bool_pn, transmit_pn = CAN_option.transmitCAN(0x602, transmit_inf,0)
        if not bool_pn:
            return False

        while True:
            bool_pn, m_can_obj = CAN_option.receiveCANbyID(0x582, 2,0)
            if bool_pn:
                break
            if bool_pn == 'stopReceive':
                return 'stopReceive', m_can_obj.Data
            elif not bool_pn:
                return False, m_can_obj.Data

        for j in range(4):
            if  m_can_obj.Data[j+4] != 0x00:
                pn_list.append(int(m_can_obj.Data[j+4]))
            else:
                flag_pn = False
                break

        if not flag_pn:
            break

    # print('pn_list = ',pn_list)
    pnCode = ''
    for asc in pn_list:
        pnCode += chr(asc)
    # print('pnCode = ', pnCode)

    # 读设备SN
    sn_list = []
    flag_sn = True
    for i in range(4):
        transmit_inf[0] = 0x40
        transmit_inf[1] = 0xf7
        transmit_inf[2] = 0x5f
        transmit_inf[3] = 0x00 + i + 1
        bool_sn, transmit_sn = CAN_option.transmitCAN(0x602, transmit_inf,0)
        if not bool_sn:
            return False

        while True:
            bool_sn, m_can_obj = CAN_option.receiveCANbyID(0x582, 2,0)
            if bool_sn:
                break
            if bool_sn == 'stopReceive':
                return 'stopReceive', m_can_obj.Data
            elif not bool_sn:
                return False, m_can_obj.Data

        for j in range(4):
            if m_can_obj.Data[j + 4] != 0x00 and m_can_obj.Data[j + 4] != 0xff:
                sn_list.append(int(m_can_obj.Data[j + 4]))
            else:
                flag_sn = False
                break

        if not flag_sn:
            break

    # print('sn_list = ', sn_list)
    snCode = ''
    for asc in sn_list:
        snCode += chr(asc)
    # print('snCode = ', snCode)

    # 读设备rev
    rev_list = []
    flag_rev = True
    for i in range(4):
        transmit_inf[0] = 0x40
        transmit_inf[1] = 0xf9
        transmit_inf[2] = 0x5f
        transmit_inf[3] = 0x00 + i + 1
        bool_rev, transmit_rev = CAN_option.transmitCAN(0x602, transmit_inf,0)
        if not bool_rev:
            return False

        while True:
            bool_rev, m_can_obj = CAN_option.receiveCANbyID(0x582, 2,0)
            if bool_rev:
                break
            if bool_rev == 'stopReceive':
                return 'stopReceive', m_can_obj.Data
            elif not bool_rev:
                return False, m_can_obj.Data

        for j in range(4):
            if m_can_obj.Data[j + 4] != 0x00:
                rev_list.append(int(m_can_obj.Data[j + 4]))
            else:
                flag_rev = False
                break

        if not flag_rev:
            break

    # print('rev_list = ', rev_list)
    revCode = ''
    for asc in rev_list:
        revCode += chr(asc)
    # print('revCode = ', revCode)

    return True,[pnCode,snCode,revCode]

# def online_thread():
#     try:
#         time_online = time.time()
#         while True:
#             QApplication.processEvents()
#             if (time.time() - time_online) * 1000 > 2000:
#                 if not mainThreadRunning():
#                     return False
#                 self.showInf(f'错误：总线初始化超时！' + self.HORIZONTAL_LINE)
#                 QMessageBox.critical(None, '错误', '总线初始化超时！请检查CAN分析仪是否正确连接', QMessageBox.Yes |
#                                      QMessageBox.No, QMessageBox.Yes)
#                 self.endOfTest()
#                 return False
#             bool_online = self.isModulesOnline()
#             # bool_online = self.isModulesOnline(currentTable)
#             if bool_online:
#                 if not mainThreadRunning():
#                     return False
#                 self.showInf(f'总线初始化成功！' + self.HORIZONTAL_LINE)
#                 break
#             else:
#                 if not mainThreadRunning():
#                     return False
#                 self.showInf(f'错误：总线初始化失败！再次尝试初始化。' + self.HORIZONTAL_LINE)
#                 # QMessageBox.critical(None, '错误', '总线初始化失败！再次尝试初始化', QMessageBox.Yes |
#                 #                  QMessageBox.No, QMessageBox.Yes)
#                 self.PassOrFail(False)
#                 if not mainThreadRunning():
#                     return False
#         self.showInf('模块在线检测结束！' + self.HORIZONTAL_LINE)
#
#     except:
#         if not mainThreadRunning():
#             return False
#         QMessageBox(QMessageBox.Critical, '错误提示', '总线初始化异常，请检查设备').exec_()
#         self.PassOrFail(False)
#         return False
