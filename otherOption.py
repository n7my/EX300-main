# -*- coding: utf-8 -*-
import CAN_option
import win32print
from ctypes import Array, c_ubyte

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


def isModulesOnline(CAN1,CAN2,module_1,module_2,waiting_time,CANAddr_relay,type:str):
    # 检测设备心跳
    if check_heartbeat( CAN1, module_1, waiting_time) == False:
        return False,0
    if type == 'AO' or type == 'AI' :
        if  check_heartbeat( CANAddr_relay, '继电器1',  waiting_time) == False:
            return False,1
        if  check_heartbeat( CANAddr_relay + 1, '继电器2', waiting_time) == False:
            return False,3
    if  check_heartbeat( CAN2, module_2, waiting_time) == False:
        return False,7
    return True,8


def check_heartbeat(can_addr, inf, max_waiting):
    can_id = 0x700 + can_addr
    bool_receive,  m_can_obj = CAN_option.receiveCANbyID(can_id, max_waiting)
    print( m_can_obj.Data)
    if bool_receive == False:
        # self.result_signal.emit(f'错误：未发现{inf}' + self.HORIZONTAL_LINE)
        return False
    # self.result_signal.emit(f'发现{inf}：收到心跳帧：{hex(self.m_can_obj.ID)}\n\n')
    return True