# -*- coding:utf-8 -*-
import numpy as np


class SimpleLinearRegressionOne(object):
    def __init__(self):
        self.a_ = None
        self.b_ = None

    def fit(self, x_train, y_train):
        assert x_train.ndim == 1, \
            "Test Data Dimension must be one"
        assert len(x_train) == len(y_train), \
            "the size of x_train must equal to the size of y_train"

        x_mean = np.mean(x_train)
        y_mean = np.mean(y_train)
        # numerator
        num = (x_train - x_mean).dot(y_train - y_mean)
        # denominator
        d = (x_train - x_mean).dot(x_train - x_mean)
        self.a_ = num / d
        self.b_ = y_mean - self.a_ * x_mean

    def predict(self, x_test):
        assert self.a_ is not None and self.b_ is not None, \
            "a and b is not None"
        print(self.a_)
        print(self.b_)
        return [self._predict(x) for x in x_test]

    def _predict(self, x_predict):
        return self.a_ * x_predict + self.b_

    def __repr__(self):
        return "SimpleLinearRegressionOne"


def main():
    r1 = SimpleLinearRegressionOne()
    x = np.array([651, 749, 835, 920, 1005, 1091, 1176, 1262, 1347])
    y = np.array([1547, 2303, 2961, 3618, 4274, 4930, 5587, 6244, 6904])
    r1.fit(x, y)
    r1.predict(np.array([1000]))


if __name__ == '__main__':
    main()