import getprovincereport
import pymsql.


if __name__ == '__main__':
    pr = getprovincereport.ProvinceReport()
    pr.get_summary_of_grade()