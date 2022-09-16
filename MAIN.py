import win32com.client as win32
import os
import shutil


class Main: 
    
    fileName = None
    buildingNum = 0
    sheetsName = None
    buildingName = None
    path = os.getcwd()
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = None
    ws = None
    locationRowTop = 3 # 페이지 내 첫번째 시작 행
    locationRowMid = 13 # 페이지 내 두번쨰 시작 행
    locationRowBottom = 23 # 페이지 내 세번째 시작 행
    location_Col_1 = "B" # 고정 열
    location_Col_2 = "U" # 번호확인 함수 위치
    location_Col_3 = "O" # 설명확인 함수 위치
    imageCycle = 1   # 각 동의 폴더에서 넣을 사진의 순서
    contentCycle =  1 # 조사표 내용이 삽입되는 순서
    fileList = None
    fileListJpg = None
    imageNum = 0  # .jpg로 끝나는 사진의 개수
    pageNum = 0 # 사진이 무조건 다 들어가는 페이지수
    lastPageNum = 0 # 마지막 페이지에서 모자란 사진 숫자

    def setFileName(self, fileName):
        self.fileName = fileName
    def getFileName(self):
        return self.fileName
  
    def setSheetsName(self, sheetsName):
        self.sheetsName = sheetsName
    def getSheetsName(self):
        return self.sheetsName
    
    def setBuildingName(self, buildingName):
        self.buildingName = buildingName
    def getBuildingName(self):
        return self.buildingName

    def setPath(self):
        self.path = os.getcwd()
    def getPath(self):
        return self.path
    
    def setExcel(self):
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
    def getExcel(self):
        return self.excel

    def setWorkBook(self):
        self.wb = self.excel.Workbooks.Open(self.path + "\\" + self.fileName + ".xlsx")
    def getWorkBook(self):
        return self.wb
    
    def setWorkSheets(self):
        self.ws = self.wb.Sheets(self.sheetsName)
    def getWorkSheets(self):
        return self.ws

    def setFileList(self):
        self.fileList = os.listdir(self.path + "\\" + str(self.buildingName))
    def getFileList(self):
        return self.fileList
    
    def setFileListJpg(self):
        self.fileListJpg = [file for file in self.fileList if file.endswith(".jpg") or file.endswith("JPG")]   
    def getFileListJpg(self):
        return self.fileListJpg
    
    def setImageNum(self):
        self.imageNum = len(self.fileListJpg)  # .jpg로 끝나는 사진의 개수
    def getImageNum(self):
        return self.imageNum
    
    def setPageNum(self):
        self.pageNum = self.imageNum // 3 # 사진이 무조건 다 들어가는 페이지수
    def getPageNum(self):
        return self.pageNum
    
    def setLastPageNum(self):
        self.lastPageNum = 3 - (self.imageNum % 3) # 마지막 페이지에서 모자란 사진 숫자
    def getLastPageNum(self):
        return self.lastPageNum


    # 이미지를 복사해주는 함수    
    def duplicateImage(self):
        if(self.imageNum % 3 > 0):
            self.pageNum = self.pageNum +1 # 다들어가지 않고 남는 사진이 있으면 남는 사진이 들어간는 페이지 수 추가
        
        
            #사진 갯수가 3의 배수가 아니면 마지막 사진을 그다음번호로 복사함 
            for k in range(1, self.lastPageNum+1, 1):
                shutil.copyfile(self.path + "\\" + str(self.buildingName) + "\\" + str(self.imageNum) + ".jpg", self.path + "\\" + str(self.buildingName) + "\\" + str(self.imageNum+k) + ".jpg") 


    # 화살표를 복사해주는 함수
    def inputArrow(self, location_Row):
        self.ws.Range("Z1:AF4").Copy(self.ws.Range("U"+str(location_Row)))
    

    # 결함의 위치를 삽입해주는 함수
    def inputLocation(self, location_Row):
        self.ws.Cells(int(location_Row)+1, 17).Value = (str(self.ws.Cells(int(self.contentCycle)+8, 33)))+"\n"+(str(self.ws.Cells(int(self.contentCycle)+8, 35)))
        return ()


    # 조사표 내용을 조사사진에 넣어주는 함수
    def inputContents(self,location_Row):
        #41,43,45,47,49    25,38
        self.ws.Cells(int(location_Row)+8, 15).Value = self.ws.Cells(int(self.contentCycle)+9, 41)
        self.ws.Cells(int(location_Row)+8, 17).Value = self.ws.Cells(int(self.contentCycle)+9, 43)
        self.ws.Cells(int(location_Row)+8, 19).Value = self.ws.Cells(int(self.contentCycle)+9, 45)
        self.ws.Cells(int(location_Row)+8, 21).Value = self.ws.Cells(int(self.contentCycle)+9, 47)
        self.ws.Cells(int(location_Row)+8, 23).Value = self.ws.Cells(int(self.contentCycle)+9, 49)
        self.ws.Cells(int(location_Row)+8, 25).Value = self.ws.Cells(int(self.contentCycle)+9, 38)
        self.contentCycle = self.contentCycle + 1
        return(self.contentCycle)


    # 설명번호와 사진번호가 맞는지 확인해주는 엑셀수식을 삽입해주는 함수
    def checkImageNum(self, location_Row):
        self.ws.Cells(int(location_Row)+1, 21).Value = "=IF(MID(O"+str(int(location_Row)+1)+",SEARCH(\".\",O"+str(int(location_Row)+1)+",1),3)=MID(W"+str(int(location_Row)+8)+",SEARCH(\".\",W"+str(int(location_Row)+8)+",1),3),\"\",\"번호확인\")" 
        self.ws.Cells(int(location_Row)+1, 21).Font.Color = -16776961
        return()


    # 설명부분 내용을 합성해주는 엑셀수식을 삽입해주는 함수
    def combineExplanation(self, location_Row, location_Col):
        if(self.ws.Cells(int(location_Row)+8, 25).Value == "균열"):
            self.ws.Cells(int(location_Row)+4, 15).Value = "="+ location_Col + str(int(location_Row) + 8) + "&"+"\"균열\"" 
        else:
            self.ws.Cells(int(location_Row)+4, 15).Value = self.ws.Cells(int(location_Row)+8, 25)
        return()


    # 이전 행(Row)을 받고 다음 사진이 삽입될 행(Row)의 넘버를 반환해주는 함수
    def nextLocation(self, location_Row):
        location_Row = location_Row + 32
        return(location_Row)


    # 행, 열, 사진경로, 폴더이름, 사진 번호를 받고 사진을 해당위치에 삽입 및 다음 사진 넘버를 반환 해주는 함수
    def inputImage(self, location_Row, location_Col):
        location = location_Col + str(location_Row)
        rng = self.ws.Range(location) 
        ImagePath = self.path+"\\" + str(self.buildingName) + "\\" + str(self.imageCycle) + ".jpg" 
        Image = self.ws.Shapes.AddPicture(ImagePath, False,True, rng.Left, rng.Top, 247.68, 184.28)
        self.imageCycle = self.imageCycle + 1
        return(self.imageCycle) 

