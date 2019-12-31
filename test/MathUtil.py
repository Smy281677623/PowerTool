# -*- coding:utf-8 -*-
import numpy as np
import math


class MathUtils(object):
    def __init__(self):
        self.a_ = None
        self.b_ = None
        self.c_ = None

    def fit(self, x_train, y_train):
        assert x_train is not None and y_train is not None, \
            "x_train, y_train can not be None"

        x = np.array(x_train)
        y = np.array(y_train)
        return self._predict(x, y)

    def _predict(self, x , y):
        f1 = np.polyfit(x, y, 2)
        return f1


def main():
    a = MathUtils()
    x = [598, 722, 798, 878, 956, 1033, 1108, 1188, 1263]
    y = [1151, 2051, 3249, 3249, 3841, 4433, 5025, 5617, 6212]
    result = a.fit(x, y)
    print(result[0])
    print(result[1])
    print(result[2])
    # c
    print(int(round(round(result[0], 8) * 1000000, 2)))
    # b
    print(int(round(result[1], 8) * 10000))
    # a
    print(int(round(result[2], 8) * 10000))


if __name__ == '__main__':
    main()