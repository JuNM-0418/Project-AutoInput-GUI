import MAIN as M
import tkinter as tk
import tkinter.messagebox
import tkinter.ttk as ttk
import win32com.client as win32
import time


def InputImage():

    buttonStart.config(state="disabled")

    Start.setFileName(entryFileName.get())

    SheetsName = str(entrySheetsName.get())
    SheetsName = SheetsName.replace(" ","")
    SheetsNameList = SheetsName.split(',')


    BuildingName = str(entryBuildingName.get())
    BuildingName = BuildingName.replace(" ","")
    BuildingNameList = BuildingName.split(',')


    WorkState = True


    for SheetsName, BuildingName in zip(SheetsNameList, BuildingNameList):
        try:
            Start.setSheetsName(SheetsName)
            Start.setBuildingName(BuildingName)

            Start.setWorkBook()
            Start.setWorkSheets()
            Start.setFileList()
            Start.setFileListJpg()
            Start.setImageNum()
            Start.setPageNum()
            Start.setLastPageNum()

            Start.DuplicateImage()

            PageNum = int(Start.getPageNum())
            
            WorkBarState = tk.DoubleVar()
            WorkBar = ttk.Progressbar(root,maximum=PageNum,variable=WorkBarState)   
            WorkBar.pack() 

        except:
            tkinter.messagebox.showerror("에러 발생", "엑셀 파일, 시트 또는 폴더 이름을 확인해주세요.")
            break

        for i in range(0, int(PageNum), 1):
            time.sleep(0.01)
            WorkBarState.set(i+1)
            WorkBar.update()
            
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
                WorkState = False
                tkinter.messagebox.showerror("에러 발생", str(Start.BuildingName) + " " + str(Start.ImageCycle) + ".jpg 사진이 없습니다.")
                break
                
            Start.CombineExplanation(Start.location_Row_1,Start.location_Col_2)
            Start.CombineExplanation(Start.location_Row_2,Start.location_Col_2)
            Start.CombineExplanation(Start.location_Row_3,Start.location_Col_2)
            Start.CheckImageNum(Start.location_Row_1)
            Start.CheckImageNum(Start.location_Row_2)
            Start.CheckImageNum(Start.location_Row_3)

            Start.location_Row_1 = Start.NextLocation(Start.location_Row_1)
            Start.location_Row_2 = Start.NextLocation(Start.location_Row_2)
            Start.location_Row_3 = Start.NextLocation(Start.location_Row_3)
            
        if(WorkState == False) :
            break


        Start.location_Row_1 = 3 # 페이지 내 첫번째 시작 행
        Start.location_Row_2 = 13 # 페이지 내 두번쨰 시작 행
        Start.location_Row_3 = 23 # 페이지 내 세번째 시작 행
        Start.location_Col_1 = "B" # 고정 열
        Start.location_Col_2 = "U" # 번호확인 함수 위치
        Start.location_Col_3 = "O" # 설명확인 함수 위치
        Start.ImageCycle = 1   # 각 동의 폴더에서 넣을 사진의 순서
        Start.ContentsCycle =  1 # 조사표 내용이 삽입되는 순서

    Start.excel.Visible=True
    tkinter.messagebox.showinfo("작업 완료", "모든 사진이 삽입 되었습니다.")
    root.destroy()
    

Start = M.Main()
root = tk.Tk()
root.title("안전점검")
root.geometry('300x300')


labelFileName = tk.Label(root, text = "파일 이름을 입력해주세요.")
labelFileName.pack()
entryFileName = tk.Entry(root)
entryFileName.pack()

labelSheetsName = tk.Label(root,text="시트 이름을 입력해주세요.")
labelSheetsName.pack()
entrySheetsName = tk.Entry(root)
entrySheetsName.pack()


labelBuildingName = tk.Label(root, text="동 이름을 입력해주세요.")
labelBuildingName.pack()
entryBuildingName = tk.Entry(root)
entryBuildingName.pack()

buttonStart = tk.Button(root,text="입력 완료", command=InputImage)
buttonStart.pack()


root.mainloop()
