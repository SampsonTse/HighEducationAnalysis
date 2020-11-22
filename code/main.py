import province
import city
import pandas as pd
#import catchdata.catch as cd

# dshs = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "12", "13", "14", "15", "16", "17", "18", "19", "20",
            # "51", "52", "53"]

if __name__ == '__main__':

    dshs = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "12", "13", "14", "15", "16", "17", "18", "19", "20",
    "51", "52", "53"]
    for dsh in dshs:
        print(dsh)
        city_p = city.city_report(dsh)
        city_p.test2()

    print("数学已更新")


    dshs = [ "17", "18", "19", "20","51", "52", "53"]

    for dsh in dshs:
        print(dsh)
        city_p = city.city_report(dsh)
        city_p.test()




















