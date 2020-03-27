import openpyxl
from tkinter import *
from tkinter import filedialog as fd
from tkinter import ttk
import formulas
import win32com.client
import os
import pathlib
from ACPRaw import ACPRaw
import csv
import time
import threading

RawDataList=list()
RawDataWithPCI=list()
root= Tk()
fileDirect= ""

def ImportRawDataCSV(name):
    with open(name) as csvfile:
        reader = csv.reader(csvfile,delimiter=',')
        for row in reader:
            if 'StIDSecID' not in row[0]:
                 RawDataList.append(ACPRaw(row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16],
                                           row[17], row[18], row[19], row[20], row[21], row[22], row[23], row[24], row[25], row[26], row[27], row[28], row[29], row[30], row[31], row[32], row[33],
                                           row[34], row[35], row[36], row[37]))
def OutputCSV():
    with open(os.path.join(fileDirect,'Raw Data PCI.csv'),'w',newline='') as csvFileOut:
        print(os.path.join(fileDirect,'Raw Data PCI.csv'))
        writer=csv.writer(csvFileOut)
        writer.writerow(['StIDSecID','Sample Number', 'Rater', 'StreetName','Begin Location', 'End Location', 'Sample Length', 'Sample Width','Date','Sample Notes','Photos','QA','Special','Sample Area','Alligator L','Alligator M','Alligator H','Block L','Block M','Block H',
                        'Distortion L','Distortion M','Distortion H','LongTrans L','LongTrans M','LongTrans H','Patch L', 'Patch M', 'Patch H','Raveling L','Raveling M','Raveling H','RuttingDepression L','RuttingDepression M','RuttingDepression H','Weathering L','Weathering M', 'Weathering H','Calculated PCI'])
        for Sample in RawDataWithPCI:
            writer.writerow([Sample.StIDSecID,Sample.SampleNumber,Sample.Rater,Sample.StreetName,Sample.BegLocation,Sample.EndLocation,Sample.SampleLength,Sample.SampleWidth,Sample.Date,Sample.SampleNotes,Sample.Photos,Sample.QA,Sample.Special,Sample.SampleArea,
                            Sample.AlligatorL,Sample.AlligatorM,Sample.AlligatorH,Sample.BlockL,Sample.BlockM,Sample.BlockH,Sample.DistortionL,Sample.DistortionM,Sample.DistortionH,Sample.LongTransL,Sample.LongTransM,Sample.LongTransH,
                            Sample.PatchL,Sample.PatchM,Sample.PatchH,Sample.RavelingL,Sample.RavelingM,Sample.RavelingH,Sample.RuttingDepressionL,Sample.RuttingDepressionM,Sample.RuttingDepressionH,Sample.WeatheringL,Sample.WeatheringM,Sample.WeatheringH, Sample.CalcPCI])

def WriteToFile(iterat,wb):
        wb['CalcMany'].cell(row=iterat+3,column=2).value=RawDataList[iterat].SampleArea
        wb['CalcMany'].cell(row=iterat+3,column=3).value=RawDataList[iterat].AlligatorL
        wb['CalcMany'].cell(row=iterat+3,column=4).value=RawDataList[iterat].AlligatorM
        wb['CalcMany'].cell(row=iterat+3,column=5).value=RawDataList[iterat].AlligatorH
        wb['CalcMany'].cell(row=iterat+3,column=6).value=RawDataList[iterat].BlockL
        wb['CalcMany'].cell(row=iterat+3,column=7).value=RawDataList[iterat].BlockM
        wb['CalcMany'].cell(row=iterat+3,column=8).value=RawDataList[iterat].BlockH
        wb['CalcMany'].cell(row=iterat+3,column=9).value=RawDataList[iterat].DistortionL
        wb['CalcMany'].cell(row=iterat+3,column=10).value=RawDataList[iterat].DistortionM
        wb['CalcMany'].cell(row=iterat+3,column=11).value=RawDataList[iterat].DistortionH
        wb['CalcMany'].cell(row=iterat+3,column=12).value=RawDataList[iterat].LongTransL
        wb['CalcMany'].cell(row=iterat+3,column=13).value=RawDataList[iterat].LongTransM
        wb['CalcMany'].cell(row=iterat+3,column=14).value=RawDataList[iterat].LongTransH
        wb['CalcMany'].cell(row=iterat+3,column=15).value=RawDataList[iterat].PatchL
        wb['CalcMany'].cell(row=iterat+3,column=16).value=RawDataList[iterat].PatchM
        wb['CalcMany'].cell(row=iterat+3,column=17).value=RawDataList[iterat].PatchH
        wb['CalcMany'].cell(row=iterat+3,column=18).value=RawDataList[iterat].RavelingL
        wb['CalcMany'].cell(row=iterat+3,column=19).value=RawDataList[iterat].RavelingM
        wb['CalcMany'].cell(row=iterat+3,column=20).value=RawDataList[iterat].RavelingH
        wb['CalcMany'].cell(row=iterat+3,column=21).value=RawDataList[iterat].RuttingDepressionL
        wb['CalcMany'].cell(row=iterat+3,column=22).value=RawDataList[iterat].RuttingDepressionM
        wb['CalcMany'].cell(row=iterat+3,column=23).value=RawDataList[iterat].RuttingDepressionH
        wb['CalcMany'].cell(row=iterat+3,column=24).value=RawDataList[iterat].WeatheringL
        wb['CalcMany'].cell(row=iterat+3,column=25).value=RawDataList[iterat].WeatheringM
        wb['CalcMany'].cell(row=iterat+3,column=26).value=RawDataList[iterat].WeatheringH

def RefreshExcel(wb):
        wb.save('PCI_ACP_Calculator.xlsx')
        xlapp = win32com.client.DispatchEx("Excel.Application")
        file = os.path.abspath('PCI_ACP_Calculator.xlsx')
        wb2 = xlapp.workbooks.Open(file)
        wb2.RefreshAll()
        wb2.Save()
        xlapp.Quit()
        
def CalculateValues():
    wb=openpyxl.load_workbook('PCI_ACP_Calculator.xlsx')
    counter =0
    StartingCount=len(RawDataList)
    while len(RawDataList) > 0:

        if len(RawDataList)>90:
            for i in range(90):
                WriteToFile(i,wb)
                counter+=1
                print('Row ' + str(counter) + " of " +str(StartingCount))

            RefreshExcel(wb)

            wb3=openpyxl.load_workbook('PCI_ACP_Calculator.xlsx',data_only=True)
            for t in range(90):
                RawDataList[t].CalcPCI=wb3['CalcMany'].cell(row=t+3,column=27).value
                RawDataWithPCI.append(RawDataList[t])
        
            for y in range(90):
                RawDataList.pop(0)

        elif len(RawDataList)<90:
            for last in range(len(RawDataList)):
                WriteToFile(last,wb)
                counter+=1
                print('Row ' + str(counter) + " of " +str(StartingCount))

            RefreshExcel(wb)

            wb3=openpyxl.load_workbook('PCI_ACP_Calculator.xlsx',data_only=True)
            for lastt in range(len(RawDataList)):
                RawDataList[lastt].CalcPCI=wb3['CalcMany'].cell(row=lastt+3,column=27).value
                RawDataWithPCI.append(RawDataList[lastt])
        
            RawDataList.clear()
        OutputCSV()
        

def callback():
    root.filename = fd.askopenfilename(initialdir="C:",title="Select Raw Data CSV", filetypes=(("csv files","*.csv"),))
    ImportRawDataCSV(root.filename)
    fileDirect = os.path.dirname(root.filename)
    thread1=threading.Thread(target=CalculateValues, args=())
    progress(thread1)

def progress(thread):
    thread.start()
    pb1= ttk.Progressbar(root,orient=HORIZONTAL, mode='indeterminate')
    pb2= ttk.Progressbar(root,orient=HORIZONTAL, mode='determinate')
    pb2['value']=100

    pb1.pack()
    pb1.start()

    while thread.is_alive():
        root.update()
        pass
    pb1.destroy()
    pb2.pack()

root.title('PCI Calculator')
root.minsize(300,100)
button= Button(root,text='Browse',width='30',command=callback)
button.pack()
root.mainloop()
