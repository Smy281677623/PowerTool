# -*- coding:utf-8 -*-
from sklearn import linear_model        #表示，可以调用sklearn中的linear_model模块进行线性回归。
import numpy as np


if __name__ == '__main__':
    model = linear_model.LinearRegression()
    X = [[651], [749], [835], [920], [1005], [1091], [1176], [1262], [1347]]
    Y = [[1547], [2303], [2961], [3618], [4274], [4930], [5587], [6244], [6904]]
    x = np.array([651, 749, 835, 920, 1005, 1091, 1176, 1262, 1347])
    y = np.array([1547, 2303, 2961, 3618, 4274, 4930, 5587, 6244, 6904])
    model.fit(X, Y)
    print(model.intercept_)  # 截距
    print(model.coef_)  # 线性模型的系数
    a = model.predict([[1000]])