from ctypes import *
import time
from ctypes import Array, c_ubyte
from threading import Thread
import datetime
from typing import Type

# from DIDOAIAOControl import Ui_Control
STATUS_OK = 1
RESERVED = 0  # 保留字段

"""1.读取动态链接库"""
# 依赖的DLL文件(存放在根目录下)
CAN_DLL_PATH = 'D:/MyData/wujun89/Downloads/EX300-main/ControlCAN.dll'
# 读取DLL文件
Can_DLL = windll.LoadLibrary(CAN_DLL_PATH)
"""2.设置初始化CAN的配置参数"""
# CAN卡类别为 USBCAN-2A, USBCAN-2C, CANalyst-II
VCI_USB_CAN_2 = 4
# CAN卡下标索引, 比如当只有一个USB-CAN适配器时, 索引号为0, 这时再插入一个USB-CAN适配器那么后面插入的这个设备索引号就是1, 以此类推
DEV_INDEX = 0
# CAN通道，默认为CAN2
CAN_INDEX = 1
# 过滤验收码
ACC_CODE = 0x80000000
# 过滤屏蔽码
ACC_MASK = 0xFFFFFFFF
# 滤波模式 0/1=接收所有类型
FILTER = 0
# 波特率 T0 默认1000k
TIMING_0 = 0x00
# 波特率 T1 默认1000k
TIMING_1 = 0x14
# 工作模式 0=正常工作
MODE = 0

receiveFlag = True
receivePauseFlag = False




"""2.VCI_OpenDevice 打开设备"""
# 打开设备, 一个设备只能打开一次
# return: 1=OK 0=ERROR

def connect(VCI_USB_CAN_2, DEV_INDEX, RESERVED):
    # VCI_USB_CAN_2: 设备类型
    # DEV_INDEX:     设备索引
    # RESERVED:      保留参数
    ret = Can_DLL.VCI_OpenDevice(VCI_USB_CAN_2, DEV_INDEX, RESERVED)
    if ret == 1:
        print('VCI_OpenDevice: 设备开启成功')
        # time.sleep(1000)
    else:
        print('VCI_OpenDevice: 设备开启失败')
        # time.sleep(1000)

    return ret


"""3.VCI_InitCAN 初始化指定CAN通道"""
# 通道初始化参数结构
# AccCode:  过滤验收码
# AccMask:  过滤屏蔽码
# Reserved: 保留字段
# Filter:   滤波模式 0/1=接收所有类型 2=只接收标准帧 3=只接收扩展帧
# Timing0:  波特率 T0
# Timing1:  波特率 T1
# Mode:     工作模式 0=正常工作 1=仅监听模式 2=自发自收测试模式
class VCI_CAN_INIT_CONFIG(Structure):
    _fields_ = [
        ("AccCode", c_uint),
        ("AccMask", c_uint),
        ("Reserved", c_uint),
        ("Filter", c_ubyte),
        ("Timing0", c_ubyte),
        ("Timing1", c_ubyte),
        ("Mode", c_ubyte)
    ]

# 过滤验收码
ACC_CODE = 0x80000000

# 过滤屏蔽码
ACC_MASK = 0xFFFFFFFF

# 滤波模式 0/1 = 接收所有类型
FILTER = 0

# 波特率 T0 = 1000k
TIMING_0 = 0x00

# 波特率 T1 = 1000k
TIMING_1 = 0x14

# 工作模式 0 = 正常工作
MODE = 0
init_config = VCI_CAN_INIT_CONFIG(ACC_CODE, ACC_MASK, RESERVED, FILTER, TIMING_0, TIMING_1, MODE)

# 初始化通道
# return: 1=OK 0=ERROR
def init(VCI_USB_CAN_2, DEV_INDEX, can_index,init_config):
    # init_config = VCI_CAN_INIT_CONFIG(ACC_CODE, ACC_MASK, RESERVED, FILTER, TIMING_0, TIMING_1, MODE)
    # VCI_USB_CAN_2: 设备类型
    # DEV_INDEX:     设备索引
    # can_index:     CAN通道索引
    # init_config:   请求参数体
    ret = Can_DLL.VCI_InitCAN(VCI_USB_CAN_2, DEV_INDEX, can_index, byref(init_config))
    if ret == 1:
        print('VCI_InitCAN: 通道 ' + str(can_index + 1) + ' 初始化成功')
    else:
        print('VCI_InitCAN: 通道 ' + str(can_index + 1) + ' 初始化失败')
    return ret


"""4.VCI_StartCAN 打开指定CAN通道"""
# 打开通道
# return: 1=OK 0=ERROR
def start(VCI_USB_CAN_2, DEV_INDEX, can_index):
    # VCI_USB_CAN_2: 设备类型
    # DEV_INDEX:     设备索引
    # can_index:     CAN通道索引
    time1 = time.time()
    while True:
        if (time.time() - time1) > 2000:
            return False
        ret = Can_DLL.VCI_StartCAN(VCI_USB_CAN_2, DEV_INDEX, can_index)
        if ret == STATUS_OK:
            print('VCI_StartCAN: 通道 ' + str(can_index + 1) + ' 打开成功')
            break
        else:
            print('VCI_StartCAN: 通道 ' + str(can_index + 1) + f' 打开失败，再次尝试打开通道{str(can_index + 1)}')

    return True


"""5.VCI_Transmit 发送数据"""


# CAN帧结构体
# ID:         帧ID, 32位变量, 数据格式为靠右对齐
# TimeStamp:  设备接收到某一帧的时间标识, 时间标示从CAN卡上电开始计时, 计时单位为0.1ms
# TimeFlag:   是否使用时间标识, 为1时TimeStamp有效, TimeFlag和TimeStamp只在此帧为接收帧时才有意义
# SendType:   发送帧类型 0=正常发送(发送失败会自动重发, 重发时间为4秒, 4秒内没有发出则取消) 1=单次发送(只发送一次, 发送失败不会自动重发, 总线只产生一帧数据)[二次开发, 建议1, 提高发送的响应速度]
# RemoteFlag: 是否是远程帧 0=数据帧 1=远程帧(数据段空)
# ExternFlag: 是否是扩展帧 0=标准帧(11位ID) 1=扩展帧(29位ID)
# DataLen:    数据长度DLC(<=8), 即CAN帧Data有几个字节, 约束了后面Data[8]中的有效字节
# Data:       CAN帧的数据, 由于CAN规定了最大是8个字节, 所以这里预留了8个字节的空间, 受DataLen约束, 如DataLen定义为3, 即Data[0]、Data[1]、Data[2]是有效的
# Reserved:   保留字段
class VCI_CAN_OBJ(Structure):
    _fields_ = [
        ("ID", c_uint),
        ("TimeStamp", c_uint),
        ("TimeFlag", c_ubyte),
        ("SendType", c_ubyte),
        ("RemoteFlag", c_ubyte),
        ("ExternFlag", c_ubyte),
        ("DataLen", c_ubyte),
        ("Data", c_ubyte * 8),
        ("Reserved", c_ubyte * 3)
    ]

# class VCI_CAN_OBJ(Structure):
#     id = c_uint( )
#     timeStamp = c_uint(0)
#     timeFlag = c_ubyte(0)
#     sendType = c_ubyte(0)
#     remoteFlag = c_ubyte(0)
#     externFlag = c_ubyte(0)
#     dataLen = c_ubyte(8)
#     date = [c_ubyte(0), c_ubyte(0), c_ubyte(0), c_ubyte(0), c_ubyte(0), c_ubyte(0), c_ubyte(0), c_ubyte(0)]
#     reserved = [c_ubyte(0), c_ubyte(0), c_ubyte(0)]

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


# 发送数据
# return: 1=OK 0=ERROR
def transmit(VCI_USB_CAN_2, DEV_INDEX, can_index, sendContent):  # TRANSMIT_DATA01~TRANSMIT_DATA08为要发送的8个字节的数据
    ubyte_array_8 = c_ubyte * 8
    DATA = ubyte_array_8(sendContent[0], sendContent[1], sendContent[2], sendContent[3], sendContent[4],
                         sendContent[5], sendContent[6], sendContent[7])
    ubyte_array_3 = c_ubyte * 3
    RESERVED_3 = ubyte_array_3(RESERVED, RESERVED, RESERVED)
    can_obj = VCI_CAN_OBJ(0X602, TIME_STAMP, TIME_FLAG, TRANSMIT_SEND_TYPE, REMOTE_FLAG, EXTERN_FLAG, DATA_LEN,
                          DATA, RESERVED_3)
    # VCI_USB_CAN_2: 设备类型
    # DEV_INDEX:     设备索引
    # can_index:     CAN通道索引
    # can_obj:       请求参数体
    # TRANSMIT_LEN:  发送的帧数量
    ret = Can_DLL.VCI_Transmit(VCI_USB_CAN_2, DEV_INDEX, can_index, byref(can_obj), TRANSMIT_LEN)

    if ret == STATUS_OK:
        print('VCI_Transmit: 通道 ' + str(can_index + 1) + ' 发送数据成功')
    else:
        print('VCI_Transmit: 通道 ' + str(can_index + 1) + ' 发送数据失败')


"""6.VCI_Receive 接收数据"""


# 接收数据
# return: 1=OK 0=ERROR
# def receive(VCI_USB_CAN_2, DEV_INDEX, can_index):
#     ubyte_array_8 = c_ubyte * 8
#     DATA = ubyte_array_8(RESERVED, RESERVED, RESERVED, RESERVED, RESERVED, RESERVED, RESERVED, RESERVED)
#     ubyte_array_3 = c_ubyte * 3
#     RESERVED_3 = ubyte_array_3(RESERVED, RESERVED, RESERVED)
#     # 参数结构参考122行
#     can_obj = VCI_CAN_OBJ(RECEIVE_ID, TIME_STAMP, TIME_FLAG, RECEIVE_SEND_TYPE, REMOTE_FLAG, EXTERN_FLAG, DATA_LEN,
#                           DATA, RESERVED_3)
#     # VCI_USB_CAN_2: 设备类型
#     # DEV_INDEX:     设备索引
#     # can_index:     CAN通道索引
#     # can_obj:       请求参数体
#     # RECEIVE_LEN:   用来接收帧结构体数组的长度
#     # WAIT_TIME:     保留参数
#     ret = Can_DLL.VCI_Receive(VCI_USB_CAN_2, DEV_INDEX, can_index, byref(can_obj), RECEIVE_LEN, WAIT_TIME)
#     while ret != STATUS_OK:
#         print('VCI_Receive: 通道 ' + str(can_index + 1) + ' 接收数据失败, 正在重试')
#         ret = Can_DLL.VCI_Receive(VCI_USB_CAN_2, DEV_INDEX, can_index, byref(can_obj), RECEIVE_LEN, WAIT_TIME)
#     else:
#         print('VCI_Receive: 通道 ' + str(can_index + 1) + ' 接收数据成功')
#         print('ID: ', can_obj.ID)
#         print('DataLen: ', can_obj.DataLen)
#         print('Data: ', list(can_obj.Data))
#     return ret

def close(VCI_USB_CAN_2, DEV_INDEX):
    Can_DLL.VCI_CloseDevice(VCI_USB_CAN_2, DEV_INDEX)
    print("VCI_CloseDevice: 设备关闭成功")

def receiveCANbyID(can_id, max_waiting):
    ubyte_array_8 = c_ubyte * 8
    DATA = ubyte_array_8(RESERVED, RESERVED, RESERVED, RESERVED, RESERVED, RESERVED, RESERVED, RESERVED)
    ubyte_array_3 = c_ubyte * 3
    RESERVED_3 = ubyte_array_3(RESERVED, RESERVED, RESERVED)
    # 参数结构参考122行
    m_can_obj = VCI_CAN_OBJ(can_id, TIME_STAMP, TIME_FLAG, RECEIVE_SEND_TYPE, REMOTE_FLAG, EXTERN_FLAG, DATA_LEN, DATA, RESERVED_3)
    #
    now_time = time.time()
    while True:
        global receiveFlag
        global receivePauseFlag

        if not receiveFlag:
            print('停止接收数据1！')
            return 'stopReceive', m_can_obj

        if receivePauseFlag:
            while True:
                if not receiveFlag:
                    print('停止接收数据2！')
                    return 'stopReceive', m_can_obj
                if not receivePauseFlag:
                    break
        else:
            if (time.time() - now_time) * 1000 > 2000:
                print('接收数据超时！')
                return False, m_can_obj
            else:
                ret = Can_DLL.VCI_Receive(VCI_USB_CAN_2, DEV_INDEX, CAN_INDEX, byref(m_can_obj), RECEIVE_LEN, WAIT_TIME)
                while ret != 1:
                    print('VCI_Receive: CAN通道 ' + str(CAN_INDEX + 1) + ' 接收数据失败, 正在重试')
                    ret = Can_DLL.VCI_Receive(VCI_USB_CAN_2, DEV_INDEX, CAN_INDEX, byref(m_can_obj), RECEIVE_LEN, WAIT_TIME)
                    # print(f'time.time() - now_time:{time.time() - now_time}')
                    if (time.time() - now_time) * 1000 > 2000:
                        print('接收数据错误且超时！')
                        return False, m_can_obj
                else:
                    # print('VCI_Receive: 通道 ' + str(CAN_INDEX + 1) + f' 接收到{m_can_obj.ID}数据')
                    if m_can_obj.ID == can_id:
                        print('VCI_Receive: CAN通道 ' + str(CAN_INDEX + 1) + f' 接收到 {hex(m_can_obj.ID)} 成功')
                        print('ID: ', hex(m_can_obj.ID))
                        # print('DataLen: ', m_can_obj.DataLen)
                        print('Data: ', hex(m_can_obj.Data[0]),hex(m_can_obj.Data[1]),hex(m_can_obj.Data[2]),hex(m_can_obj.Data[3]),hex(m_can_obj.Data[4]),hex(m_can_obj.Data[5]),hex(m_can_obj.Data[6]),hex(m_can_obj.Data[7]))
                        return True, m_can_obj

def transmitCAN(can_id, sendContent):
    Can_DLL.VCI_ClearBuffer(VCI_USB_CAN_2, DEV_INDEX, CAN_INDEX)
    bool_result = True
    loopCount = 3
    # if len(sendContent) == 8:
    ubyte_array_8 = c_ubyte * 8
    DATA = ubyte_array_8(sendContent[0], sendContent[1], sendContent[2], sendContent[3],
                         sendContent[4], sendContent[5], sendContent[6], sendContent[7])
    # elif len(sendContent) == 3:
    #     ubyte_array_8 = c_ubyte * 3
    #     DATA = ubyte_array_8(sendContent[0], sendContent[1], sendContent[2])#, sendContent[3], sendContent[4], sendConte
    ubyte_array_3: Type[Array[c_ubyte]] = c_ubyte * 3
    RESERVED_3 = ubyte_array_3(RESERVED, RESERVED, RESERVED)
    can_obj = VCI_CAN_OBJ(can_id, TIME_STAMP, TIME_FLAG, TRANSMIT_SEND_TYPE, REMOTE_FLAG, EXTERN_FLAG, DATA_LEN, DATA, RESERVED_3)

    while True:
        if bool_result == False:
            time.sleep(0.5)
            Can_DLL.VCI_ClearBuffer(VCI_USB_CAN_2, DEV_INDEX, CAN_INDEX)
        bool_result = (TRANSMIT_LEN == Can_DLL.VCI_Transmit(VCI_USB_CAN_2, DEV_INDEX, CAN_INDEX, byref(can_obj), TRANSMIT_LEN))
        # Can_DLL.VCI_Transmit(VCI_USB_CAN_2, DEV_INDEX, CAN_INDEX, m_can_obj, TRANSMIT_LEN)
        loopCount = loopCount - 1
        time.sleep(0.1)

        if (bool_result == True) or loopCount <= 0:
            break
    return bool_result,can_obj


def transmitCANAddr(can_id, sendContent):
    can_id = 0x00000000
    Can_DLL.VCI_ClearBuffer(VCI_USB_CAN_2, DEV_INDEX, CAN_INDEX)
    bool_result = True
    loopCount = 5
    # if len(sendContent) == 8:
    ubyte_array_8 = c_ubyte * 8
    DATA = ubyte_array_8(sendContent[0], sendContent[1], sendContent[2], sendContent[3],
                        sendContent[4], sendContent[5], sendContent[6], sendContent[7])

    ubyte_array_3: Type[Array[c_ubyte]] = c_ubyte * 3
    RESERVED_3 = ubyte_array_3(RESERVED, RESERVED, RESERVED)
    can_obj = VCI_CAN_OBJ(can_id, TIME_STAMP, TIME_FLAG, TRANSMIT_SEND_TYPE, REMOTE_FLAG, EXTERN_FLAG, DATA_LEN, DATA, RESERVED_3)

    while True:
        if bool_result == False:
            time.sleep(0.5)
        bool_result = (TRANSMIT_LEN == Can_DLL.VCI_Transmit(VCI_USB_CAN_2, DEV_INDEX, CAN_INDEX, byref(can_obj), TRANSMIT_LEN))

        loopCount = loopCount - 1
        time.sleep(0.1)

        if (bool_result == True) or loopCount <= 0:
            break
    return bool_result,can_obj

def receiveStop():
    global receiveFlag
    receiveFlag = False

def receiveRun():
    global receiveFlag
    receiveFlag = True

def receivePause():
    global receivePauseFlag
    receivePauseFlag = True

def receiveResume():
    global receivePauseFlag
    receivePauseFlag = False

