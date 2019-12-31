# -*- coding:utf-8 -*-
import numpy as np
import matplotlib.pyplot as plt

if __name__ == '__main__':
    # 定义x、y散点坐标
    x = [651, 749, 835, 920, 1005, 1091, 1176, 1262, 1347]
    x = np.array(x)
    print('x is :\n', x)
    num = [1547, 2303, 2961, 3618, 4274, 4930, 5587, 6244, 6904]
    y = np.array(num)
    print('y is :\n', y)
    # 用2次多项式拟合
    f1 = np.polyfit(x, y, 2)
    print('f1 is :\n', f1)
    p1 = np.poly1d(f1)
    print('p1 is :\n', p1)
    # 也可使用yvals=np.polyval(f1, x)
    # 拟合y值
    yvals = p1(x)
    print('yvals is :\n', yvals)
    # 绘图
    plot1 = plt.plot(x, y, 's', label='original values')
    plot2 = plt.plot(x, yvals, 'r', label='polyfit values')
    plt.xlabel('x')
    plt.ylabel('y')
    plt.legend(loc=4)
    plt.title('polyfitting')
    plt.show()
