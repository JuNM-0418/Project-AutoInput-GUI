import MAIN as M
import tkinter as tk
import win32com.client as win32
import os
import time
import shutil
from tqdm import tqdm
from tqdm import trange

def Process():
    def InputImage():
        print("Qwer")
        SheetsName = entrySheetsName.get()
        BuildingName = entryBuildinName.get()
        Start.setSheetsName(SheetsName)
        Start.setBuildingName(BuildingName)
        print(Start.getSheetsName())
        print(Start.getBuildingName())
        Start.setWorkBook()
        Start.setWorkSheets()
        Start.setFileList()
        Start.setFileListJpg()
        Start.setImageNum()
        Start.setPageNum()
        Start.setLastPageNum()
        PageNum = int(Start.getPageNum())

        Start.DuplicateImage()
        for i in range(0, int(PageNum), 1):
            try:
                # 사진, 화살표 삽입 및 조사표 내용 삽입
                Start.ImageCycle = Start.InputImage(Start.location_Row_1,Start.location_Col_1)
                Start.ContentsCycle = Start.InputContents(Start.location_Row_1)
                Start.InputArrow(Start.location_Row_1)
                Start.ImageCycle = Start.InputImage(Start.location_Row_2,Start.location_Col_1)
                Start.ContentsCycle = Start.InputContents(Start.location_Row_2)
                Start.InputArrow(Start.location_Row_2)
                Start.ImageCycle = Start.InputImage(Start.location_Row_3,Start.location_Col_1)
                Start.ContentsCycle = Start.InputContents(Start.location_Row_3)
                Start.InputArrow(Start.location_Row_3)
            except:
                print(str(Start.BuildingName) + " " + str(Start.ImageCycle) + ".jpg 사진이 없습니다.")
                
            Start.CombineExplanation(Start.location_Row_1,Start.location_Col_2)
            Start.CombineExplanation(Start.location_Row_2,Start.location_Col_2)
            Start.CombineExplanation(Start.location_Row_3,Start.location_Col_2)
            Start.CheckImageNum(Start.location_Row_1)
            Start.CheckImageNum(Start.location_Row_2)
            Start.CheckImageNum(Start.location_Row_3)


            Start.location_Row_1 = Start.NextLocation(Start.location_Row_1)
            Start.location_Row_2 = Start.NextLocation(Start.location_Row_2)
            Start.location_Row_3 = Start.NextLocation(Start.location_Row_3)

            print(Start.BuildingName + " 사진 삽입 완료.")

        print("모든 사진 삽입 완료.")

        Start.excel.Visible=True

    print("ASdf")
    FileName = entryFileName.get()
    BuildingNum = int(entryBuildingNum.get())
    Start.setFileName(FileName)
    Start.setBuildingNum(BuildingNum)
    print(Start.getFileName())
    print(Start.getBuildingNum())

    
    labelSheetsName = tk.Label(root,text="시트 이름을 입력해주세요.")
    labelSheetsName.pack()
    entrySheetsName = tk.Entry(root)
    entrySheetsName.pack()

    labelBuildingName = tk.Label(root, text="동 이름을 입력해주세요.")
    labelBuildingName.pack()
    entryBuildinName = tk.Entry(root)
    entryBuildinName.pack()

    buttonStart = tk.Button(root,text="입력 완료", command=InputImage)
    buttonStart.pack()

Start = M.Main()
root = tk.Tk()

labelFileName = tk.Label(root, text = "파일 이름을 입력해주세요.")
labelFileName.pack()
entryFileName = tk.Entry(root)
entryFileName.pack()

labelBuildingNum = tk.Label(root, text = "동의 수를 입력해주세요.")
labelBuildingNum.pack()
entryBuildingNum = tk.Entry(root)
entryBuildingNum.pack()

buttonNext = tk.Button(root, text = "다음",command=Process)
buttonNext.pack()





root.mainloop()
