import MAIN as Document
import tkinter as tk
import tkinter.messagebox
import tkinter.ttk as ttk
from tkinter import filedialog
import time


def selectFile():
    labelSelectSheet.pack()
    buttonSelectSheet.pack()


    excelFilePath = filedialog.askopenfilename(initialdir='./',title='파일선택', filetypes=(('xlsx files','*.xlsx'),('all files','*.*')))
    Start.setExcelFilePath(excelFilePath) 
    Start.setWorkBook()
    Start.setWorkSheetsName()
    buttonSelectFile.config(state="disabled")
   
    

def selectSheet():
    buttonSelectSheet.config(state="disabled")

    listbox = tk.Listbox(root, selectmode='extended', height=5)
    for sheet in enumerate(Start.getWorkSheetsName()):
        listbox.insert(sheet[0], sheet[1])
    listbox.pack()


    def selectComplete():
        selectedSheetList = listbox.curselection()

        for selectedSheet in selectedSheetList:
            Start.setSheetsNameList(Start.getWorkSheetsName()[selectedSheet])

        listbox.delete(0, len(Start.getWorkSheetsName()))
        for index in range(len(selectedSheetList)):
            listbox.insert(index, Start.getSheetsNameList()[index])

        listbox.config(state="disabled")
        buttonComplete.config(state="disabled")
        createImgPathButton()

    buttonComplete = tk.Button(root, text="선택완료",command = selectComplete)
    buttonComplete.pack()
    tkinter.ttk.Separator(root, orient="horizontal").pack(fill="both")



def createImgPathButton():
    index = 1
    for size in range(len(Start.getSheetsNameList())):
        labelListSelectImgPath.append(tk.Label(root, text=f"사진 폴더{index}를 선택하세요."))
        labelListSelectImgPath[size].pack()
        buttonListSelectImgPath.append(tk.Button(root, text="열기", command=selectImgFilePath))
        buttonListSelectImgPath[size].pack()
        index+=1

    
    
    
def selectImgFilePath():
    global check
    check+=1
    imgFilePath = filedialog.askdirectory(initialdir="./",title='폴더 선택')
    Start.setImgFilePathList(imgFilePath)
    

    if (check == len(Start.getSheetsNameList())):
        for index in range(len(Start.getSheetsNameList())):
            buttonListSelectImgPath[index].config(state="disabled")
        tkinter.ttk.Separator(root, orient="horizontal").pack(fill="both")

        buttonInsertImg.pack()


def inputImg():
    buttonInsertImg.config(state="disabled")
    Start.setImgFileList()
    Start.setImageNum()
    Start.setPageNum()
    Start.setLastPageNum()
    Start.setWorkSheets()
    Start.duplicateImage()
    

    
   


    for index, data in enumerate(Start.getPageNum()):

        WorkState = True
        WorkBarState = tk.DoubleVar()
        WorkBar = ttk.Progressbar(root,maximum=len(Start.getPageNum()),variable=WorkBarState)   
        WorkBar.pack()

        for i in range(0, data, 1):
            time.sleep(0.01)
            WorkBarState.set(i+1)
            WorkBar.update()

            
            Start.inputImage(Start.locationRowTop, Start.location_Col_1, index)
            Start.inputContents(Start.locationRowTop, index)
            Start.inputArrow(Start.locationRowTop, index)
            Start.inputLocation(Start.locationRowTop, index)

            Start.inputImage(Start.locationRowMid,Start.location_Col_1, index)
            Start.inputContents(Start.locationRowMid, index)
            Start.inputArrow(Start.locationRowMid, index)
            Start.inputLocation(Start.locationRowMid, index)

            Start.inputImage(Start.locationRowBottom,Start.location_Col_1, index)
            Start.inputContents(Start.locationRowBottom, index)
            Start.inputArrow(Start.locationRowBottom, index)
            Start.inputLocation(Start.locationRowBottom, index)

            Start.combineExplanation(Start.locationRowTop,Start.location_Col_2, index)
            Start.combineExplanation(Start.locationRowMid,Start.location_Col_2, index)
            Start.combineExplanation(Start.locationRowBottom,Start.location_Col_2, index)

            Start.checkImageNum(Start.locationRowTop, index)
            Start.checkImageNum(Start.locationRowMid, index)
            Start.checkImageNum(Start.locationRowBottom, index)

            Start.locationRowTop = Start.nextLocation(Start.locationRowTop)
            Start.locationRowMid = Start.nextLocation(Start.locationRowMid)
            Start.locationRowBottom = Start.nextLocation(Start.locationRowBottom)


        
 
        Start.locationRowTop = 3 # 페이지 내 첫번째 시작 행
        Start.locationRowMid = 13 # 페이지 내 두번쨰 시작 행
        Start.locationRowBottom = 23 # 페이지 내 세번째 시작 행
        Start.imageCycle = 1   # 각 동의 폴더에서 넣을 사진의 순서
        Start.contentCycle =  1 # 조사표 내용이 삽입되는 순서

    tkinter.messagebox.showinfo("작업 완료", "모든 사진이 삽입 되었습니다.")
    root.destroy()
    Start.excel.Visible=True    



Start = Document.ExcelFileInfo()

root = tk.Tk()
root.title("안전점검")
root.geometry('300x800')

global check
check = 0


excelFilePath = None
labelImgPath = []
labelListSelectImgPath = []
buttonListSelectImgPath = []

labelSelectFile = tk.Label(root, text = "파일을 선택하세요.")
labelSelectFile.pack()
buttonSelectFile = tk.Button(root, text="열기", command=selectFile)
buttonSelectFile.pack()

tkinter.ttk.Separator(root, orient="horizontal").pack(fill="both")


labelSelectSheet = tk.Label(root,text="시트를 선택하세요.")
# labelSelectSheet.pack()
buttonSelectSheet = tk.Button(root, text="열기", command=selectSheet)
# buttonSelectSheet.pack()


buttonInsertImg = tk.Button(root, text="삽입 시작", command=inputImg)

root.mainloop()
