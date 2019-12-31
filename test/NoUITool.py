# -*- coding:utf-8 -*-
from MultimeterUtil import *
from ComSerialUtil import *


def main():
    test_ser = comm_util('COM17', 115200, 8, 1, 'N')
    mm = Multimeter(('192.168.2.20', 5025))
    for i in range(0, 5):
        test_ser.read_bin_data()
    voltage_ad = []
    test_voltage = []
    for i in range(2, 11):
        data = '10 01 ' + '0' + str(hex(i))[2:] + ' de 13 0A'
        test_ser.send_hex_data(data)
        voltage_ad.append(test_ser.read_bin_data().replace('\n', ''))
        test_voltage.append(mm.GetDcVolt())
    print('电压测试AD值和实测电压值测试')
    for i in voltage_ad:
        print(i)
    for j in test_voltage:
        print(j)


if __name__ == '__main__':
    main()