# -*- coding:utf-8 -*-
import serial
from time import *


class comm_util(object):
    def __init__(self, port,bortnate,bytesize,stopbits,parity):
        # open Serial Com
        self.s = serial.Serial(port, bortnate)
        # Set bytesize
        self.s.bytesize = bytesize
        # Set parity
        self.s.parity = parity
        # Set stop bit
        self.s.stopbits = stopbits
        # Set TimeOut
        self.s.timeout = 5.0

    """读取一行数据"""
    def _readline(self, EOL = serial.CR):
        """read a line which is terminated with end-of-line (eol) character
        ('\n' by default) or until timeout."""
        eol = EOL
        leneol = len(eol)
        line = bytearray()
        while True:
            c = self.s.read(1)
            if c:
                line += c
                if line[-leneol:] == eol:
                    break
            else:
                break
        return bytes(line)

    """打开串口"""
    def open_com(self):
        self.s.open()
        return True

    """关闭串口"""
    def close_com(self):
        try:
            self.s.close()
        except Exception:
            print("Close Com Error")

    """发送数据"""
    def send_hex_data(self, data):
        try:
            self.s.write(bytes.fromhex(data))
        except Exception:
            print("发送数据异常")


    """读取数据"""
    def read_bin_data(self):
        recv_data = self._readline()
        # print(recv_data)
        return recv_data.decode('utf-8')

    def read_bin_data2(self):
        recv_data = self._readline()
        # print(recv_data)
        return recv_data

    def recv_data_line(self):
        print(self.s.readline(100))


def main():
    c = comm_util('COM17', 115200, 8, 1, 'N')
    sleep(6)
    begin_flag = c.recv_data_line()
    print(begin_flag)
    flag = 100
    print(c.s.isOpen())
    while(flag):
        c.send_hex_data('10 01 02 de 13 0A')
        sleep(5)
        recv_data = c.read_bin_data()
        flag -= 1
    c.close_com()


if __name__ == '__main__':
    main()
