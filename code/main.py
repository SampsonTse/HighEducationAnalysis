import province
import city


if __name__ == '__main__':
   # dshs = ["01","02","03","04","05","06","07","08","09","12","13","14","15","16","17","18","19","20","51","52","53"]
   dshs = ["02"]
   for dsh in dshs:
      print(dsh)
      # city_r = city.city_report(dsh)
      # city_r.ztgk()
      # city_r.dtfx()

      city_a = city.city_report_appendix(dsh)
      city_a.ysfgk()
      city_a.dtfx()









