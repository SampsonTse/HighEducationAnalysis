# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import pymysql
import ProvinceReport as pr
import numpy as np

# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    db = pymysql.connect("localhost","root","1234","higheducation")
    cursor = db.cursor()

    print(pr.LQGK(cursor,1))



# See PyCharm help at https://www.jetbrains.com/help/pycharm/
