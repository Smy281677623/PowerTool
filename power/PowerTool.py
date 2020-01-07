# -*- coding:utf-8 -*-
import sys
import serial.tools.list_ports
from ComSerialUtil import *
from PowerUI import *
from PyQt5.QtWidgets import QWidget, QApplication, QMessageBox
from PyQt5.QtCore import (QThread, pyqtSignal, pyqtSlot)
from MultimeterUtil import *


class ToolBoard(QThread):
    signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.tool_ser = serial.Serial()
        self.ser_flag = 0
        self.init_flag = False
        self.test_flag = None

    def on_source(self, com_msg):
        print(com_msg+'0000')
        self.tool_ser.port = com_msg.split(":")[0]
        self.tool_ser.baudrate = int(com_msg.split(":")[1])
        self.tool_ser.bytesize = int(com_msg.split(":")[2])
        self.tool_ser.stopbits = int(com_msg.split(":")[3])
        if com_msg.split(":")[4] is 'None':
            self.tool_ser.parity = 'N'
        try:
            self.tool_ser.open()
        except Exception:
            print('工具板串口打开失败')

    def close(self):
        self.tool_ser.close()
        return self.tool_ser.isOpen()

    @pyqtSlot()
    def run(self):
        if self.test_flag == '大电流测试':
            if self.init_flag is True:
                if self.ser_flag >= 0:
                    data = '3D 00 01 '
                    if self.ser_flag <= 15:
                        data = '3D 00 00 '+'0'+hex(self.ser_flag)[2:]+' 00 0D 0A'
                    if self.ser_flag > 15:
                        data = '3D 00 00 ' + hex(self.ser_flag)[2:] + ' 00 0D 0A'
                    sleep(0.05)
                    self.tool_ser.write(bytes.fromhex(data))
                    print(data)
                    print(self.ser_flag)
            if self.ser_flag >= 51:
                sleep(1)
                close_cmd = '3D 01 00 FF 00 0D 0A'
                self.tool_ser.write(bytes.fromhex(close_cmd))
                return
            self.ser_flag += 1
        if self.test_flag == '小电流测试':
            if self.init_flag is True:
                if self.ser_flag >= 0:
                    data = '3A 00 01 '
                    if self.ser_flag <= 15:
                        data = '3A 00 00 '+'0'+hex(self.ser_flag)[2:]+' 00 0D 0A'
                    if self.ser_flag > 15:
                        data = '3A 00 00 ' + hex(self.ser_flag)[2:] + ' 00 0D 0A'
                    sleep(0.05)
                    self.tool_ser.write(bytes.fromhex(data))
                    print(data)
                    print(self.ser_flag)
            if self.ser_flag >= 51:
                sleep(1)
                close_cmd = '3A 01 00 FF 00 0D 0A'
                self.tool_ser.write(bytes.fromhex(close_cmd))
                return
            self.ser_flag += 1



class MyThread(QThread):
    # 自定义型号，执行run()函数时，从相关线程发射此信号
    signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.source_txt = None
        self.test_what =None

    def on_source(self, ip):
        self.soc = Multimeter((ip.split(" ")[0], int(ip.split(" ")[1])))

    @pyqtSlot()
    def run(self):
        if self.test_what == 'Test Voltage':
            # 发出信号
            if self.soc.con_flag is None:
                return None
            try:
                self.source_txt = self.soc.GetDcVolt()
                self.signal.emit(str(self.source_txt))
            except Exception:
                return None
        elif self.test_what == 'Test electric_current':
            # 发出信号
            if self.soc.con_flag is None:
                return None
            try:
                self.source_txt = self.soc.GetDcCurrent()
                self.signal.emit(str(self.source_txt))
            except Exception:
                return None


class Power_Utils(QWidget, Ui_Form):
    sig = pyqtSignal(str)
    sig2 = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.com_util = None
        self.mul = None
        self.setupUi(self)
        self.bang()
        # open Serial Com
        self.ser = serial.Serial()
        # 创建线程
        self.my_thread = None
        self.tool_board_thread = ToolBoard()

    def open(self):
        if self.pushButton.text() == '打开串口':
            port = self.comboBox.currentText()
            baudrate = self.comboBox_2.currentText()
            databit = self.comboBox_3.currentText()
            stopbit = self.comboBox_4.currentText()
            parity = self.comboBox_5.currentText()
            try:
                self.ser = serial.Serial(port, int(baudrate))
                self.ser.bytesize = int(databit)
                self.ser.stopbits = int(stopbit)
                self.ser.timeout = 12.0
                if parity == 'None':
                    self.ser.parity = 'N'
                if self.ser.isOpen() is True:
                    for i in range(0, 4):
                        recv_data = self.read_bin_data()
                        self.textEdit.append(recv_data)
                    self.pushButton.setText('关闭串口')
                    self.pushButton.setStyleSheet("background: green; color: black;")
                    return
            except Exception:
                QMessageBox.warning(self, '测试板串口异常', '请检查串口是否初始化', QMessageBox.Ok)
                return
        elif self.pushButton.text() == '关闭串口':
            self.ser.close()
            if self.ser.isOpen() is False:
                self.pushButton.setText('打开串口')
                self.pushButton.setStyleSheet("background: red; color: black;")
                return

    def bang(self):
        # 串口按钮
        self.pushButton.clicked.connect(self.open)
        self.pushButton.setStyleSheet("background: red ; color: black")
        # 电压测试按钮
        self.pushButton_2.clicked.connect(self.ADvalue_Voltage)
        # 清除按钮
        self.pushButton_3.clicked.connect(self.clear_recv)
        # 串口初始化
        self.pushButton_4.clicked.connect(self.port_check)
        self.pushButton_4.setStyleSheet("background: yellow ; color: black")
        # 测试电流工具板串口
        self.pushButton_7.clicked.connect(self.open_tool_board_serial)
        self.pushButton_7.setStyleSheet("background: red ; color: black")
        # 测试板恒压设置
        self.pushButton_11.clicked.connect(self.set_constant_voltage)
        # 大电流测试
        self.pushButton_9.clicked.connect(self.test_electric_current)
        # 小电流测试
        self.pushButton_5.clicked.connect(self.test_small_current)

    def clear_recv(self):
        self.textEdit.clear()
        self.textEdit_2.clear()
        self.textEdit_3.clear()

    def ADvalue_Voltage(self):
        if self.ser.isOpen():
            self.my_thread = MyThread()
            # 将自定义信号sig连接到MyThread.on_source函数
            self.sig.connect(self.my_thread.on_source)
            # 向MyThread.on_source函数传递ip
            self.sig.emit(self.lineEdit_3.text()+' '+self.lineEdit_4.text())
            # 将自定义信号signal连接到information()槽函数
            self.my_thread.signal.connect(self.information)
            # 启动线程
            self.my_thread.start()
            self.my_thread.test_what = 'Test Voltage'
            if self.my_thread.soc.con_flag is False :
                QMessageBox.warning(self, '万用表IP不对', '万用表ip地址不对或者冲突', QMessageBox.Ok)
                return
            self.pushButton_2.setEnabled(False)
            if self.ser.isOpen():
                for i in range(2, 11):
                    data = '10 01 ' + '0' + str(hex(i))[2:] + ' de 13 0A'
                    print(data)
                    self.send_hex_data(data)
                    recv_data = self.read_bin_data()
                    self.textEdit.append(recv_data)
                    if i >= 3:
                        self.my_thread.run()
                    QApplication.processEvents()
            self.my_thread.soc.close()
            self.my_thread = None
            self.pushButton_2.setEnabled(True)
        else:
            QMessageBox.warning(self, '串口异常', '串口没打开', QMessageBox.Ok)

    """读取一行数据"""
    def _readline(self, EOL=serial.CR):
        """read a line which is terminated with end-of-line (eol) character
        ('\n' by default) or until timeout."""
        eol = EOL
        leneol = len(eol)
        line = bytearray()
        while True:
            c = self.ser.read(1)
            if c:
                line += c
                if line[-leneol:] == eol:
                    break
            else:
                break
        return bytes(line)

    """发送数据"""
    def send_hex_data(self, data):
        try:
            self.ser.write(bytes.fromhex(data))
        except Exception:
            print("发送数据异常")

    """读取数据"""
    def read_bin_data(self):
        try:
            recv_data = self._readline()
            print(recv_data[1:-1].decode('utf-8'))
            return recv_data[1:-1].decode('utf-8')
        except Exception:
            print("接受数据异常")

    # 串口检测
    def port_check(self):
        # 检测所有存在的串口，将信息存储在字典中
        self.Com_Dict = {}
        port_list = list(serial.tools.list_ports.comports())
        self.comboBox.clear()
        self.comboBox_11.clear()
        i = 1
        for port in port_list:
            self.Com_Dict["%s" % port[0]] = "%s" % port[1]
            self.comboBox.addItem(port[0])
            self.comboBox_11.addItem(port[0])
            print(port[0])
            self.comboBox.setItemText(i, port[0])
            self.comboBox_11.setItemText(i, port[0])
            i += 1
        if len(self.Com_Dict) == 0:
            QMessageBox.warning(self, '没有串口', '没有检测到任何串口', QMessageBox.Ok)

    @pyqtSlot(str)
    def information(self, info):
        if self.my_thread.test_what == 'Test Voltage':
            self.textEdit_2.append(info[1:])
        elif self.my_thread.test_what == 'Test electric_current':
            self.textEdit_3.append(str(float(info)*1000))

    def open_tool_board_serial(self):
        if self.pushButton_7.text() == '打开串口':
            msg = ''
            msg += self.comboBox_11.currentText()+':'
            msg += self.comboBox_12.currentText()+':'
            msg += self.comboBox_13.currentText()+':'
            msg += self.comboBox_14.currentText()+':'
            msg += self.comboBox_15.currentText()
            self.sig2.connect(self.tool_board_thread.on_source)
            self.sig2.emit(msg)
            if self.tool_board_thread.tool_ser.isOpen() is False:
                QMessageBox.warning(self, '工具板串口异常', '工具板串口打开异常,请串口初始化后选择对应的串口连接', QMessageBox.Ok)
                return
            self.tool_board_thread.init_flag = False
            self.tool_board_thread.start()
            self.pushButton_7.setText('关闭串口')
            self.pushButton_7.setStyleSheet("background: green; color: black;")
        elif self.pushButton_7.text() == '关闭串口':
            if self.tool_board_thread.close() is False:
                self.pushButton_7.setText('打开串口')
                self.pushButton_7.setStyleSheet("background: red; color: black;")

    def set_constant_voltage(self):
        """恒压值设置,测试工具板的串口必须打开"""
        if self.ser.isOpen():
            command = '10 01 '+'0'+str(hex(int(self.comboBox_16.currentText()))[2:])+' 13 0A'
            self.send_hex_data(command)
            self.read_bin_data()
            # 测试电流按钮在恒压值设置之后才能打开
            self.pushButton_11.setEnabled(False)
        else:
            QMessageBox.warning(self, '测试板串口异常', '请检查测试板串口是否打开', QMessageBox.Ok)

    def test_electric_current(self):
        """测试大电流触发按钮"""
        if self.ser.isOpen() and self.tool_board_thread.tool_ser.isOpen():
            if self.pushButton_11.isEnabled() is False:
                self.pushButton_9.setEnabled(False)
                self.my_thread = MyThread()
                self.tool_board_thread.test_flag = '大电流测试'
                self.sig.connect(self.my_thread.on_source)
                self.sig.emit(self.lineEdit_3.text() + ' ' + self.lineEdit_4.text())
                self.my_thread.signal.connect(self.information)
                if self.my_thread.soc.con_flag is False:
                    QMessageBox.warning(self, '万用表IP不对', '万用表ip地址不对或者冲突', QMessageBox.Ok)
                    return
                data = '10 01 ' + '0' + str(hex(int(self.comboBox_16.currentText()))[2:]) + ' 13 0A'
                # AD值
                # 小电流测试
                self.send_hex_data('10 01 21 de 13 0A')
                self.read_bin_data()
                self.tool_board_thread.init_flag = True
                self.my_thread.start()
                self.my_thread.test_what = 'Test electric_current'
                for i in range(1, 52):
                    self.tool_board_thread.run()
                    self.send_hex_data(data)
                    recv_data = self.read_bin_data()
                    self.textEdit.append(recv_data)
                    self.my_thread.run()
                    QApplication.processEvents()
                self.tool_board_thread.run()
                self.tool_board_thread.test_flag = None
                self.tool_board_thread.ser_flag = 0
                self.my_thread.soc.close()
                self.my_thread = None
                self.pushButton_11.setEnabled(True)
                self.pushButton_9.setEnabled(True)
            else:
                QMessageBox.warning(self, '档位没有设置', '请检查档位是否打开', QMessageBox.Ok)
        else:
            QMessageBox.warning(self, '通行异常', '请检查测试板串口和工具板串口是否打开', QMessageBox.Ok)

    def test_small_current(self):
        """测试小电流按钮触发"""
        if self.ser.isOpen() and self.tool_board_thread.tool_ser.isOpen():
            if self.pushButton_11.isEnabled() is False:
                self.pushButton_5.setEnabled(False)
                self.my_thread = MyThread()
                self.tool_board_thread.test_flag = '小电流测试'
                self.sig.connect(self.my_thread.on_source)
                self.sig.emit(self.lineEdit_3.text() + ' ' + self.lineEdit_4.text())
                self.my_thread.signal.connect(self.information)
                if self.my_thread.soc.con_flag is False:
                    QMessageBox.warning(self, '万用表IP不对', '万用表ip地址不对或者冲突', QMessageBox.Ok)
                    return
                data = '10 01 ' + '0' + str(hex(int(self.comboBox_16.currentText()))[2:]) + ' 13 0A'
                # 小电流测试
                self.send_hex_data('10 01 21 de 13 0A')
                self.read_bin_data()
                self.tool_board_thread.init_flag = True
                self.my_thread.start()
                self.my_thread.test_what = 'Test electric_current'
                for i in range(1, 52):
                    self.tool_board_thread.run()
                    self.send_hex_data(data)
                    recv_data = self.read_bin_data()
                    self.textEdit.append(recv_data)
                    self.my_thread.run()
                    QApplication.processEvents()
                self.tool_board_thread.run()
                self.tool_board_thread.test_flag = None
                self.tool_board_thread.ser_flag = 0
                self.my_thread.soc.close()
                self.my_thread = None
                self.pushButton_11.setEnabled(True)
                self.pushButton_5.setEnabled(True)
            else:
                QMessageBox.warning(self, '档位没有设置', '请检查档位是否打开', QMessageBox.Ok)
        else:
            QMessageBox.warning(self, '通行异常', '请检查测试板串口和工具板串口是否打开', QMessageBox.Ok)

def main():
    app = QApplication(sys.argv)
    power = Power_Utils()
    power.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
