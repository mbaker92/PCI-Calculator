import openpyxl
from tkinter import *
from tkinter import filedialog as fd
import formulas
import win32com.client
import os
import pathlib
from ACPRaw import ACPRaw
import csv
import time

RawDataList=list()
RawDataWithPCI=list()
root= Tk()
def ImportRawDataCSV(name):
    with open(name) as csvfile:
        reader = csv.reader(csvfile,delimiter=',')
        for row in reader:
            if 'StIDSecID' not in row[0]:
                 RawDataList.append(ACPRaw(row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16],
                                           row[17], row[18], row[19], row[20], row[21], row[22], row[23], row[24], row[25], row[26], row[27], row[28], row[29], row[30], row[31], row[32], row[33],
                                           row[34], row[35], row[36], row[37]))

def CalculateValues():
    wb=openpyxl.load_workbook('PCI_ACP_Calculator.xlsx')
    
    #while len(RawDataList) > 90:
    for i in range(90):
        wb['CalcMany'].cell(row=i+3,column=2).value=RawDataList[i].SampleArea
        wb['CalcMany'].cell(row=i+3,column=3).value=RawDataList[i].AlligatorL
        wb['CalcMany'].cell(row=i+3,column=4).value=RawDataList[i].AlligatorM
        wb['CalcMany'].cell(row=i+3,column=5).value=RawDataList[i].AlligatorH
        wb['CalcMany'].cell(row=i+3,column=6).value=RawDataList[i].BlockL
        wb['CalcMany'].cell(row=i+3,column=7).value=RawDataList[i].BlockM
        wb['CalcMany'].cell(row=i+3,column=8).value=RawDataList[i].BlockH
        wb['CalcMany'].cell(row=i+3,column=9).value=RawDataList[i].DistortionL
        wb['CalcMany'].cell(row=i+3,column=10).value=RawDataList[i].DistortionM
        wb['CalcMany'].cell(row=i+3,column=11).value=RawDataList[i].DistortionH
        wb['CalcMany'].cell(row=i+3,column=12).value=RawDataList[i].LongTransL
        wb['CalcMany'].cell(row=i+3,column=13).value=RawDataList[i].LongTransM
        wb['CalcMany'].cell(row=i+3,column=14).value=RawDataList[i].LongTransH
        wb['CalcMany'].cell(row=i+3,column=15).value=RawDataList[i].PatchL
        wb['CalcMany'].cell(row=i+3,column=16).value=RawDataList[i].PatchM
        wb['CalcMany'].cell(row=i+3,column=17).value=RawDataList[i].PatchH
        wb['CalcMany'].cell(row=i+3,column=18).value=RawDataList[i].RavelingL
        wb['CalcMany'].cell(row=i+3,column=19).value=RawDataList[i].RavelingM
        wb['CalcMany'].cell(row=i+3,column=20).value=RawDataList[i].RavelingH
        wb['CalcMany'].cell(row=i+3,column=21).value=RawDataList[i].RuttingDepressionL
        wb['CalcMany'].cell(row=i+3,column=22).value=RawDataList[i].RuttingDepressionM
        wb['CalcMany'].cell(row=i+3,column=23).value=RawDataList[i].RuttingDepressionH
        wb['CalcMany'].cell(row=i+3,column=24).value=RawDataList[i].WeatheringL
        wb['CalcMany'].cell(row=i+3,column=25).value=RawDataList[i].WeatheringM
        wb['CalcMany'].cell(row=i+3,column=26).value=RawDataList[i].WeatheringH

    wb.save('PCI_ACP_Calculator.xlsx')
    xlapp = win32com.client.DispatchEx("Excel.Application")
    file = os.path.abspath('PCI_ACP_Calculator.xlsx')
    wb2 = xlapp.workbooks.Open(file)
    wb2.RefreshAll()
    wb2.Save()
    xlapp.Quit()
        
    wb3=openpyxl.load_workbook('PCI_ACP_Calculator.xlsx',data_only=True)
    for t in range(90):
        RawDataList[t].CalcPCI=wb3['CalcMany'].cell(row=t+3,column=27).value
        RawDataWithPCI.append(RawDataList[t])
        print(RawDataWithPCI[t].CalcPCI)
        

def callback():
    root.filename = fd.askopenfilename(initialdir="C:",title="Select Raw Data CSV", filetypes=(("csv files","*.csv"),))
    ImportRawDataCSV(root.filename)
    CalculateValues()
    print(root.filename)

root.title('PCI Calculator')
root.minsize(300,100)
button= Button(root,text='Browse',width='30',command=callback)
button.pack()
root.mainloop()