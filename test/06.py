# -*- coding:utf-8 -*-
def main():
    info = [7.123499949291086e-07, 7.681010484360738, -3472.2277849876195]
    print(str(int(round(info[2], 8) * 10000)))
    print(str(int(round(info[1], 8) * 10000)))
    print(str(int(round(round(info[0], 8) * 1000000, 2))))


if __name__ == '__main__':
    main()