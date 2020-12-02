import province
import city
import pandas as pd
import catchdata.catch as cd

dshs = [ "02", "03", "04", "05", "06", "07", "08", "09", "12", "13", "14", "15", "16", "17", "18", "19", "20",
            "51", "52", "53"]

if __name__ == '__main__':

    print("英语")
    dshs = [  "12", "13", "14", "15", "16", "17", "18", "19", "20",
            "51", "52", "53"]
    for dsh in dshs:
        print(dsh)
        cp = city.city_report(dsh)
        cp.test()

    print("数学")
    dshs = [ "08", "09", "12", "13", "14", "15", "16", "17", "18", "19", "20",
            "51", "52", "53"]
    for dsh in dshs:
        print(dsh)
        cp = city.city_report(dsh)
        cp.test2()


















