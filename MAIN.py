import win32com.client as win32
import os
import time
import shutil



class Main: 
    
    FileName = None
    BuildingNum = 0
    SheetsName = None
    BuildingName = None
    Path = os.getcwd()
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = None
    ws = None
    location_Row_1 = 3 # 페이지 내 첫번째 시작 행
    location_Row_2 = 13 # 페이지 내 두번쨰 시작 행
    location_Row_3 = 23 # 페이지 내 세번째 시작 행
    location_Col_1 = "B" # 고정 열
    location_Col_2 = "U" # 번호확인 함수 위치
    location_Col_3 = "O" # 설명확인 함수 위치
    ImageCycle = 1   # 각 동의 폴더에서 넣을 사진의 순서
    ContentsCycle =  1 # 조사표 내용이 삽입되는 순서
    FileList = None
    FileListJpg = None
    ImageNum = 0  # .jpg로 끝나는 사진의 개수
    PageNum = 0 # 사진이 무조건 다 들어가는 페이지수
    LastPageNum = 0 # 마지막 페이지에서 모자란 사진 숫자

    def setFileName(self, FileName):
        self.FileName = FileName
    def getFileName(self):
        return self.FileName

    def setBuildingNum(self,BuildingNum):
        self.BuildingNum = BuildingNum
    def getBuildingNum(self):
        return self.BuildingNum
    
    def setSheetsName(self, SheetsName):
        self.SheetsName = SheetsName
    def getSheetsName(self):
        return self.SheetsName
    
    def setBuildingName(self, BuildingName):
        self.BuildingName = BuildingName
    def getBuildingName(self):
        return self.BuildingName

    def setPath(self):
        self.Path = os.getcwd()
    def getPath(self):
        return self.Path
    
    def setExcel(self):
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
    def getExcel(self):
        return self.excel

    def setWorkBook(self):
        self.wb = self.excel.Workbooks.Open(self.Path + "\\" + self.FileName + ".xlsx")
    def getWorkBook(self):
        return self.wb
    
    def setWorkSheets(self):
        self.ws = self.wb.Sheets(self.SheetsName)
    def getWorkSheets(self):
        return self.ws

    def setFileList(self):
        self.FileList = os.listdir(self.Path + "\\" + str(self.BuildingName))
    def getFileList(self):
        return self.FileList
    
    def setFileListJpg(self):
        self.FileListJpg = [file for file in self.FileList if file.endswith(".jpg") or file.endswith("JPG")]   
    def getFileListJpg(self):
        return self.FileListJpg
    
    def setImageNum(self):
        self.ImageNum = len(self.FileListJpg)  # .jpg로 끝나는 사진의 개수
    def getImageNum(self):
        return self.ImageNum
    
    def setPageNum(self):
        self.PageNum = self.ImageNum // 3 # 사진이 무조건 다 들어가는 페이지수
    def getPageNum(self):
        return self.PageNum
    
    def setLastPageNum(self):
        self.LastPageNum = 3 - (self.ImageNum % 3) # 마지막 페이지에서 모자란 사진 숫자
    def getLastPageNum(self):
        return self.LastPageNum


    # 이미지를 복사해주는 함수    
    def DuplicateImage(self):
        if(self.ImageNum % 3 > 0):
            self.PageNum = self.PageNum +1 # 다들어가지 않고 남는 사진이 있으면 남는 사진이 들어간는 페이지 수 추가
        
        
            #사진 갯수가 3의 배수가 아니면 마지막 사진을 그다음번호로 복사함 
            for k in range(1, self.LastPageNum+1, 1):
                shutil.copyfile(self.Path + "\\" + str(self.BuildingName) + "\\" + str(self.ImageNum) + ".jpg", self.Path + "\\" + str(self.BuildingName) + "\\" + str(self.ImageNum+k) + ".jpg") 

    # 화살표를 복사해주는 함수
    def InputArrow(self, location_Row):
        self.ws.Range("Z1:AF4").Copy(self.ws.Range("U"+str(location_Row)))

    # 조사표 내용을 조사사진에 넣어주는 함수
    def InputContents(self,location_Row):
        #41,43,45,47,49    25,38
        self.ws.Cells(int(location_Row)+8, 15).Value = self.ws.Cells(int(self.ContentsCycle)+9, 41)
        self.ws.Cells(int(location_Row)+8, 17).Value = self.ws.Cells(int(self.ContentsCycle)+9, 43)
        self.ws.Cells(int(location_Row)+8, 19).Value = self.ws.Cells(int(self.ContentsCycle)+9, 45)
        self.ws.Cells(int(location_Row)+8, 21).Value = self.ws.Cells(int(self.ContentsCycle)+9, 47)
        self.ws.Cells(int(location_Row)+8, 23).Value = self.ws.Cells(int(self.ContentsCycle)+9, 49)
        self.ws.Cells(int(location_Row)+8, 25).Value = self.ws.Cells(int(self.ContentsCycle)+9, 38)
        self.ContentsCycle = self.ContentsCycle + 1
        return(self.ContentsCycle)

    # 설명번호와 사진번호가 맞는지 확인해주는 엑셀수식을 삽입해주는 함수
    def CheckImageNum(self, location_Row):
        self.ws.Cells(int(location_Row)+1, 21).Value = "=IF(MID(O"+str(int(location_Row)+1)+",SEARCH(\".\",O"+str(int(location_Row)+1)+",1),3)=MID(W"+str(int(location_Row)+8)+",SEARCH(\".\",W"+str(int(location_Row)+8)+",1),3),\"\",\"번호확인\")" 
        self.ws.Cells(int(location_Row)+1, 21).Font.Color = -16776961
        return()


    # 설명부분 내용을 합성해주는 엑셀수식을 삽입해주는 함수
    def CombineExplanation(self, location_Row, location_Col):
        if(self.ws.Cells(int(location_Row)+8, 25).Value == "균열"):
            self.ws.Cells(int(location_Row)+4, 15).Value = "="+ location_Col + str(int(location_Row) + 8) + "&"+"\"균열\"" 
        else:
            self.ws.Cells(int(location_Row)+4, 15).Value = self.ws.Cells(int(location_Row)+8, 25)
        return()


    # 이전 행(Row)을 받고 다음 사진이 삽입될 행(Row)의 넘버를 반환해주는 함수
    def NextLocation(self, location_Row):
        location_Row = location_Row + 32
        return(location_Row)


    # 행, 열, 사진경로, 폴더이름, 사진 번호를 받고 사진을 해당위치에 삽입 및 다음 사진 넘버를 반환 해주는 함수
    def InputImage(self, location_Row, location_Col):
        location = location_Col + str(location_Row)
        rng = self.ws.Range(location) 
        ImagePath = self.Path+"\\" + str(self.BuildingName) + "\\" + str(self.ImageCycle) + ".jpg" 
        Image = self.ws.Shapes.AddPicture(ImagePath, False,True, rng.Left, rng.Top, 247.68, 184.28)
        self.ImageCycle = self.ImageCycle + 1
        return(self.ImageCycle) 

