# -*- coding:utf-8 -*-
import openpyxl
from openpyxl.styles import Font, Border, Side, PatternFill, colors, Alignment
import time
# 4个数据集合 传进封装类 实体

class Excel_TemplateTwo(object):
    def __init__(self, board_voltage, mu_voltage, board_current, mu_current, board_num):
        self.board_voltage_ = board_voltage
        self.mu_voltage_ = mu_voltage
        self.board_current_ = board_current
        self.mu_current_ = mu_current
        self.board_num_ = board_num
        self.setup_excel()

    def setup_excel(self):
        """创建excel模版"""
        wb = openpyxl.Workbook()
        # 如果不给sheet名称则直接创建默认的sheet1名称
        if self.board_num_ is None:
            ws = wb.create_sheet('sheet1', 0)
        else:
            ws = wb.create_sheet(str(self.board_num_), 0)
        # 居中格式设置
        ali = Alignment(horizontal='center', vertical='center')
        # 边框设置
        left, right, top, bottom = [Side(style='thin', color='000000')] * 4
        myborder = Border(left=left, right=right, top=top, bottom=bottom)

        # 5种颜色配置
        SpringGreen2 = PatternFill("solid", fgColor="6A5ACD")
        DeepSkyBlue = PatternFill("solid", fgColor="00BFFF")
        Turquoise = PatternFill("solid", fgColor="40E0D0")
        Yellow3 = PatternFill("solid", fgColor="CDCD00")
        PaleGoldenrod = PatternFill("solid", fgColor="EEE8AA")
        white = PatternFill("solid", fgColor="FFFFFF")

        # 设置 1 2 3 4 5 6 列的宽度
        for col in ['A', 'B', 'C', 'D', 'E', 'F']:
            ws.column_dimensions[col].width = 20.0
        # 设置小电流模式
        for t in ['A1', 'B1', 'C1', 'D1', 'E1', 'F1']:
            ws[t].alignment = ali
            ws[t].border = myborder
            ws[t].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)

        ws['A1'].value = '大电流模式'
        ws['B1'].value = '机器编码'

        ws.merge_cells('B1:C1')
        ws['D1'].value = '板卡条码'+self.board_num_
        ws.merge_cells('D1:F1')

        ws['A2'].value = '设置电压'
        ws['B2'].value = '大电流模式采样次数'
        ws['C2'].value = '8001电压读数（mV）'
        ws['D2'].value = '万用表电压读数（mV）'
        ws['E2'].value = '8001电流读数(mA)'
        ws['F2'].value = '万用表电流读数(mA)'

        for t in ['A2', 'B2', 'C2', 'D2', 'E2', 'F2']:
            ws[t].alignment = ali
            ws[t].border = myborder
            ws[t].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)

        # 第一轮老化测试
        ws['A3'].value = '第一轮老化测试'
        ws['A3'].alignment = ali
        ws['A3'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in ['A3', 'B3', 'C3', 'D3', 'E3', 'F3']:
            ws[t].border = myborder
            ws[t].fill = white
        ws.merge_cells('A3:F3')

        ws['A4'].value = '3.6V'
        ws['A4'].alignment = ali
        ws['A4'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in (4, 9):
            ws['A'+str(t)].border = myborder
            ws['A' + str(t)].fill = SpringGreen2
            ws['B' + str(t) ].fill = SpringGreen2
            ws['B' + str(t)].border = myborder
        ws.merge_cells('A4:A8')

        ws['A9'].value = '3.8V'
        ws['A9'].alignment = ali
        ws['A9'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in (9, 14):
            ws['A' + str(t)].border = myborder
            ws['A' + str(t)].fill = DeepSkyBlue
            ws['B' + str(t)].fill = DeepSkyBlue
            ws['B' + str(t)].border = myborder
        ws.merge_cells('A9:A13')

        ws['A14'].value = '4.0V'
        ws['A14'].alignment = ali
        ws['A14'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in (14, 19):
            ws['A' + str(t)].border = myborder
            ws['A' + str(t)].fill = Turquoise
            ws['B' + str(t)].fill = Turquoise
            ws['B' + str(t)].border = myborder
        ws.merge_cells('A14:A19')

        ws['A19'].value = '4.1V'
        ws['A19'].alignment = ali
        ws['A19'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in (19, 24):
            ws['A' + str(t)].border = myborder
            ws['A' + str(t)].fill = Yellow3
            ws['B' + str(t)].fill = Yellow3
            ws['B' + str(t)].border = myborder
        ws.merge_cells('A19:A23')

        ws['A24'].value = '4.2V'
        ws['A24'].alignment = ali
        ws['A24'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in (24, 29):
            ws['A' + str(t)].border = myborder
            ws['A' + str(t)].fill = PaleGoldenrod
            ws['B' + str(t)].fill = PaleGoldenrod
            ws['B' + str(t)].border = myborder
        ws.merge_cells('A24:A28')

        # 第二轮老化测试
        ws['A29'].value = '第二轮老化测试'
        ws['A29'].alignment = ali
        ws['A29'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in ['A29', 'B29', 'C29', 'D29', 'E29', 'F29']:
            ws[t].border = myborder
            ws[t].fill = white
        ws.merge_cells('A29:F29')

        ws['A30'].value = '3.6V'
        ws['A30'].alignment = ali
        ws['A30'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in (30, 35):
            ws['A' + str(t)].border = myborder
            ws['A' + str(t)].fill = SpringGreen2
            ws['B'+ str(t)].fill = SpringGreen2
            ws['B' + str(t)].border = myborder
        ws.merge_cells('A30:A34')

        ws['A35'].value = '3.8V'
        ws['A35'].alignment = ali
        ws['A35'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in (35, 40):
            ws['A' + str(t)].border = myborder
            ws['A' + str(t)].fill = DeepSkyBlue
            ws['B' + str(t)].fill = DeepSkyBlue
            ws['B' + str(t)].border = myborder
        ws.merge_cells('A35:A39')

        ws['A40'].value = '4.0V'
        ws['A40'].alignment = ali
        ws['A40'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in (40, 45):
            ws['A' + str(t)].border = myborder
            ws['A' + str(t)].fill = Turquoise
            ws['B' + str(t)].fill = Turquoise
            ws['B' + str(t)].border = myborder
        ws.merge_cells('A40:A44')

        ws['A45'].value = '4.1V'
        ws['A45'].alignment = ali
        ws['A45'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in (45, 50):
            ws['A' + str(t)].border = myborder
            ws['A' + str(t)].fill = Yellow3
            ws['B' + str(t)].fill = Yellow3
            ws['B' + str(t)].border = myborder
        ws.merge_cells('A45:A49')

        ws['A50'].value = '4.2V'
        ws['A50'].alignment = ali
        ws['A50'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in (50, 55):
            ws['A' + str(t)].border = myborder
            ws['A' + str(t)].fill = PaleGoldenrod
            ws['B' + str(t)].fill = PaleGoldenrod
            ws['B' + str(t)].border = myborder
        ws.merge_cells('A50:A54')

        # 第三轮老化测试
        ws['A55'].value = '第三轮老化测试'
        ws['A55'].alignment = ali
        ws['A55'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in ['A55', 'B55', 'C55', 'D55', 'E55', 'F55']:
            ws[t].border = myborder
            ws[t].fill = white
        ws.merge_cells('A55:F55')

        ws['A56'].value = '3.6V'
        ws['A56'].alignment = ali
        ws['A56'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in (56, 61):
            ws['A' + str(t)].border = myborder
            ws['A' + str(t)].fill = SpringGreen2
            ws['B' + str(t)].fill = SpringGreen2
            ws['B' + str(t)].border = myborder
        ws.merge_cells('A56:A60')

        ws['A61'].value = '3.8V'
        ws['A61'].alignment = ali
        ws['A61'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in (61, 66):
            ws['A' + str(t)].border = myborder
            ws['A' + str(t)].fill = DeepSkyBlue
            ws['B' + str(t)].fill = DeepSkyBlue
            ws['B' + str(t)].border = myborder
        ws.merge_cells('A61:A65')

        ws['A66'].value = '4.0V'
        ws['A66'].alignment = ali
        ws['A66'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in (66, 71):
            ws['A' + str(t)].border = myborder
            ws['A' + str(t)].fill = Turquoise
            ws['B' + str(t)].fill = Turquoise
            ws['B' + str(t)].border = myborder
        ws.merge_cells('A66:A71')

        ws['A71'].value = '4.1V'
        ws['A71'].alignment = ali
        ws['A71'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in (71, 76):
            ws['A' + str(t)].border = myborder
            ws['A' + str(t)].fill = Yellow3
            ws['B' + str(t)].fill = Yellow3
            ws['B' + str(t)].border = myborder
        ws.merge_cells('A71:A75')

        ws['A76'].value = '4.2V'
        ws['A76'].alignment = ali
        ws['A76'].font = Font(name='黑体', size=11, bold=False, italic=True, color=colors.BLACK)
        for t in (76, 81):
            ws['A' + str(t)].border = myborder
            ws['A' + str(t)].fill = PaleGoldenrod
            ws['B' + str(t)].fill = PaleGoldenrod
            ws['B' + str(t)].border = myborder
        ws.merge_cells('A76:A80')

        q = [q for q in range(1, 6)]
        color_array = [SpringGreen2,SpringGreen2,SpringGreen2,SpringGreen2,SpringGreen2,
                       DeepSkyBlue,DeepSkyBlue,DeepSkyBlue,DeepSkyBlue,
                       Turquoise,Turquoise,Turquoise,Turquoise,Turquoise,
                       Yellow3,Yellow3,Yellow3,Yellow3,Yellow3,
                       PaleGoldenrod,PaleGoldenrod,PaleGoldenrod,PaleGoldenrod,PaleGoldenrod]

        for i,j,h in zip(range(4, 29), q*5, color_array):
            ws.cell(i, 2, j)
            ws.cell(i,2).fill = h
            ws.cell(i,2).border = myborder

        for i, j,h in zip(range(30, 55), q * 5, color_array):
            ws.cell(i, 2, j)
            ws.cell(i, 2).fill = h
            ws.cell(i, 2).border = myborder

        for i, j,h in zip(range(56, 81), q * 5, color_array):
            ws.cell(i, 2, j)
            ws.cell(i, 2).fill = h
            ws.cell(i, 2).border = myborder

        # 从这里开始开始数据添加
        board_voltage_one = self.board_voltage_[0:25]
        board_voltage_two = self.board_voltage_[25:50]
        board_voltage_three = self.board_voltage_[50:]

        board_current_one = self.board_current_[0:25]
        board_current_two = self.board_current_[25:50]
        board_current_three = self.board_current_[50:]

        mu_voltage_one = self.mu_voltage_[0:25]
        mu_voltage_two = self.mu_voltage_[25:50]
        mu_voltage_three = self.mu_voltage_[50:]

        mu_current_one = self.mu_current_[0:25]
        mu_current_two = self.mu_current_[25:50]
        mu_current_three = self.mu_current_[50:]

        # 第一组数据添加
        for j,i in zip(range(4, 29), board_voltage_one):
            ws.cell(j, 3, i)

        for j,i in zip(range(4, 29),mu_voltage_one):
            ws.cell(j, 4, i)

        for j, i in zip(range(4, 29), board_current_one):
            ws.cell(j, 5, i)

        for j, i in zip(range(4, 29), mu_current_one):
            ws.cell(j, 6, i)

        # 第二组数据添加
        for j,i in zip(range(30, 55), board_voltage_two):
            ws.cell(j, 3, i)

        for j,i in zip(range(30, 55),mu_voltage_two):
            ws.cell(j, 4, i)

        for j, i in zip(range(30, 55), board_current_two):
            ws.cell(j, 5, i)

        for j, i in zip(range(30, 55), mu_current_two):
            ws.cell(j, 6, i)


        # 第三组数据添加
        for j,i in zip(range(56, 81), board_voltage_three):
            ws.cell(j, 3, i)

        for j,i in zip(range(56, 81), mu_voltage_three):
            ws.cell(j, 4, i)

        for j, i in zip(range(56, 81), board_current_three):
            ws.cell(j, 5, i)

        for j, i in zip(range(56, 81), mu_current_three):
            ws.cell(j, 6, i)

        prev_msg = str(time.localtime().tm_year) + '_' + str(time.localtime().tm_mon) + '_' + str(
            time.localtime().tm_mday) + '_' + str(time.localtime().tm_hour) + '_' + str(
            time.localtime().tm_min) + '_' + str(
            time.localtime().tm_sec)
        if self.board_num_ is None:
            middle_msg = 'sheet1'
        else:
            middle_msg = str(self.board_num_)
        next_msg = '100R'
        msg = prev_msg+'_'+middle_msg+'_'+next_msg
        wb.save('./'+msg+'.xlsx')


def main():
    e = Excel_TemplateTwo([1],[1], [1], [1], '1')
    # print((1 and 1 and 1 and 1 and 1) is False)

if __name__ == '__main__':
    main()