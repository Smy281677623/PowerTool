# -*- coding:utf-8 -*-
import socket

class Multimeter(object):
    def __init__(self, ip=None):
        self.s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.s.settimeout(2.0)
        self.ip = ip
        self.handle = 0
        self.rhandle = 0
        self.slen = 0
        self.rlen = 0
        self.cmd = 0
        self.subcmd = 0
        self.rcmd = 0
        self.rslt = 0
        self.con_flag = True
        try:
            if self.ip is not None:
                self.s.connect(self.ip)
        except Exception:
            self.con_flag = False

    def __del__(self):
        self.s.close()

    def connect(self, ip):
        self.ip = ip
        self.s.connect(self.ip)
        return self

    def close(self):
        self.s.close()

    def response(self, cmd):
        return ((~cmd & 0x80) | (cmd & 0x7f));

    def StringCmd(self, string_cmd):
        databytes_send = bytes(string_cmd, encoding='utf-8')
        self.s.send(databytes_send)
        return self._recv_line()

    def _recv(self,nlen):
        _data = bytearray()
        while len(_data) < nlen:
            _data += self.s.recv(nlen-len(_data))
        return _data

    def _recv_line(self,eol='\n'):
        eol = bytes(eol,encoding='utf-8')
        leneol = len(eol)
        line = bytearray()
        while True:
            c = self._recv(1)
            if c:
                line += c
                if line[-leneol:] == eol:
                    break
            else:
                break
        return bytes(line)

    def GetDcVolt(self):
        # print("MEASure:VOLTage:DC? DEF,DEF\n")
        volt = float(self.StringCmd("MEASure:VOLTage:DC? DEF,DEF\n").decode())
        # print("actual volt is", volt)
        return volt

    def GetDcCurrent(self):
        # print("MEASure:CURRent:DC? DEF,DEF\n")
        cur = float(self.StringCmd("MEASure:CURRent:DC? DEF,DEF\n").decode())
        # print(cur)
        return cur


def main():
    mm = Multimeter(("192.168.2.20", 5025))
    for i in range(0, 10):
        print(mm.GetDcVolt())


if __name__ == '__main__':
    main()