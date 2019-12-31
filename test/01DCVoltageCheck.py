# -*- coding:utf-8 -*-
from MultimeterUtil import *
from ComSerialUtil import *


def main():
    test_ser = comm_util('COM17', 115200, 8, 1, 'N')
    mm = Multimeter(('192.168.2.20', 5025))
    for i in range(0, 5):
        test_ser.read_bin_data()
    voltage_board = []
    multimeter_voltage = []
    x = [3.6, 3.8, 4.0, 4.1, 4.2]
    for i in x:
        if i <= 4.0:
            command = '0' + hex(int(i * 1000))[2:]
        else:
            command = hex(int(i * 1000))[2:]
        test_ser.send_hex_data('10 01 11 00 00 '+command[0:2]+' '+command[2:]+' 13 0A')
        sleep(0.05)
        test_ser.read_bin_data2()
        sleep(0.05)
        test_ser.send_hex_data('10 01 42 de 13 0a')
        sleep(0.05)
        voltage_board.append(eval('0x'+test_ser.read_bin_data2().hex()[18:22]))
        multimeter_voltage.append(mm.GetDcVolt()*-1000)
    print('板子反馈数据')
    print(voltage_board)
    print('万用表读数据')
    print(multimeter_voltage)


if __name__ == '__main__':
    main()