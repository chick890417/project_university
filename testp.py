import numpy as np
import pandas as pd
import statistics as sta
from matplotlib import pyplot as plt
file1 = 'Future_20180206_I020'
file2 = 'Option_20180206_I020'
data = pd.read_csv(file2+".csv")
#in order to read excel
prod_num = []
for prod_name in pd.unique(data["PROD"]):
    prod_num.append(prod_name)
#print(prod_num)
#print(len(prod_num))
def de_mean(x):
  x_bar = sta.mean(x)
  return [x_i - x_bar for x_i in x]
def dot(x, y):
    dot_product = sum(v_i * w_i for v_i, w_i in zip(x, y))
    dot_product /= (len(x))
    return dot_product
def correlation(x, y):
    dot_xy = dot(de_mean(x), de_mean(y))
    sd_x=sta.stdev(x)
    sd_y=sta.stdev(y)
    return dot_xy/(sd_x*sd_y)
index=0
writer = pd.ExcelWriter(file2+'_correlation.xlsx')
for index in range(len(prod_num)-1):
    df = pd.read_excel(r'Option_20180206_I020_5min.xlsx',sheet_name=index)
    #x=sta.stdev(df["open"])
    #y=sta.stdev(df["high"])
    #z=sta.stdev(df["low"])
    #w=sta.stdev(df["close"])
    xy=correlation(df["open"],df["high"])
    xz=correlation(df["open"],df["low"])
    xw=correlation(df["open"],df["close"])
    yz=correlation(df["high"],df["low"])
    yw=correlation(df["high"],df["close"])
    zw=correlation(df["low"],df["close"])
    data = {'open_high_cor':[xy],
            'open_low_cor':[xz],
            'open_close_cor':[xw],
            'high_low_cor':[yz],
            'high_close_cor':[yw],
            'low_close_cor':[zw]
    }
    df = pd.DataFrame(data)
    df.to_excel(writer,sheet_name = prod_num[index])
    writer.save()
    #print(xy)

writer.close()




  
