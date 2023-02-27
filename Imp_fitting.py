import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from scipy.optimize import curve_fit

address="./data/1228.xlsx"  #1228대신 해당 파일이름
re=pd.DataFrame()
def func(x, r1, r2, c1, c2):
    return -(((6.28) * x * r1 * r1 * c1) / (1 + ((6.28) * x * r1 * c1) ** 2)) - (((6.28) * x * r2 * r2 * c2) / (1 + ((6.28) * x * r2 * c2) ** 2))
for a in range(2,17):

    file=pd.read_excel(address,sheet_name=a,usecols='A:C')
    xls = pd.ExcelFile(address)
    sheets = xls.sheet_names
    current=sheets[a]

    #############################################################################################################

    x=file['Freq']
    y=file['X']
    # plt.plot(x,y)

    # plt.show()
    popt, pcov=curve_fit(func, x,y)
    if popt[0]>popt[1]:
        data=pd.DataFrame({'r1':popt[0],'r2':popt[1],'c1':popt[2],'c2':popt[3]},index=[current])
    else:
        data =pd.DataFrame ({'r1': popt[1], 'r2': popt[0], 'c1': popt[3], 'c2': popt[2]},index=[current])
    re=pd.concat([re,data])
    plt.plot(x,y, label='original')
    plt.plot(x, func(x, *popt), label='fitting')
    plt.legend()
    plt.xscale('symlog')
    plt.title('Current : '+current)
    # plt.show()
    plt.savefig("./res/png/"+current+'.png',format='png')
    plt.cla()


for a in range(17,30):
    file=pd.read_excel(address,sheet_name=a,usecols='A:C')
    xls = pd.ExcelFile(address)
    sheets = xls.sheet_names
    current=sheets[a]

    #############################################################################################################

    x=file['Freq']
    y=file['X']
    x2 = x[40:]
    y2 = y[40:]
    # plt.plot(x,y)

    # plt.show()
    popt, pcov=curve_fit(func, x2,y2)
    if popt[0]>popt[1]:
        data=pd.DataFrame({'r1':popt[0],'r2':popt[1],'c1':popt[2],'c2':popt[3]},index=[current])
    else:
        data =pd.DataFrame ({'r1': popt[1], 'r2': popt[0], 'c1': popt[3], 'c2': popt[2]},index=[current])
    re=pd.concat([re,data])
    plt.plot(x,y, label='original')
    plt.plot(x2, func(x2, *popt), label='fitting')
    plt.legend()
    plt.xscale('symlog')
    plt.title('Current : '+current)
    # plt.show()
    plt.savefig("./res/png/"+current+'.png',format='png')
    plt.cla()

print(re)
re.to_excel('./res/xlsx/fitting.xlsx')