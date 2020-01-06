# -*- coding:utf-8 -*-
from MultimeterUtil import *
from ComSerialUtil import *
from ExcelTemplate import *


def main():
    borad_num = str(input('请输入板卡条码'))
    test_ser = comm_util('COM17', 115200, 8, 1, 'N')
    tool_ser = comm_util('COM7', 9600, 8, 1, 'N')
    board_voltage = []
    board_current = []
    mu_voltage = []
    mu_current = []
    mm = Multimeter(('192.168.2.20', 5025))
    for i in range(0, 5):
        test_ser.read_bin_data()
    for j in range(1, 4):
        print('第'+str(j)+'组数据')
        x = [3.6, 3.8, 4.0, 4.1, 4.2]
        for i in x:
            print('电压--------'+str(i*1000)+'----------情况')
            if i <= 4.0:
                command = '0' + hex(int(i * 1000))[2:]
            else:
                command = hex(int(i * 1000))[2:]
            # 设置电压
            test_ser.send_hex_data('10 01 11 00 00 ' + command[0:2] + ' ' + command[2:] + ' 13 0A')
            sleep(0.05)
            test_ser.read_bin_data2()
            sleep(0.05)
            for x in range(1, 11):
                tool_ser.send_hex_data('3A 00 00 0'+hex(x)[2:]+' 00 0D 0A')
                sleep(0.05)
                # 读取板子反馈的电压电流值
                test_ser.send_hex_data('10 01 42 de 13 0a')
                sleep(0.05)
                temp_data = test_ser.s.read(28).hex()
                # print(temp_data)
                print(str(eval('0x'+temp_data[18:22]))+'\t', end="")
                board_voltage.append(str(eval('0x'+temp_data[18:22])))
                print(str(mm.GetDcVolt() * -1000) + '\t', end="")
                mu_voltage.append(str(mm.GetDcVolt() * -1000))
                sleep(0.05)
                # 万用表读取的电压电流值
                print(str(eval('0x' + temp_data[10:14]) / 10) + '\t', end="")
                board_current.append(str(eval('0x' + temp_data[10:14]) / 10))
                sleep(0.05)
                print(str(mm.GetDcCurrent()*1000)+'')
                mu_current.append(str(mm.GetDcCurrent()*1000))
            tool_ser.send_hex_data('3A 00 00 FF 00 0D 0A')
    tool_ser.s.close()
    test_ser.s.close()
    mm.close()
    print(board_voltage)
    print(mu_voltage)
    print(board_current)
    print(mu_current)
    Excel_Template(board_voltage,mu_voltage,board_current,mu_current,borad_num)


if __name__ == '__main__':
    main()

