import win32com.client as win32
import os
import shutil


class Info: 
    
    excelFilePath = None
    imgFilePathList = []
    sheetsNameList = []
    imgFileList = []
    imageNum = []  # .jpg로 끝나는 사진의 개수
    pageNum = [] # 사진이 무조건 다 들어가는 페이지수
    lastPageNum = [] # 마지막 페이지에서 모자란 사진 숫자

   
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = None
    ws = []
    wsName = None
    locationRowTop = 3 # 페이지 내 첫번째 시작 행
    locationRowMid = 13 # 페이지 내 두번쨰 시작 행
    locationRowBottom = 23 # 페이지 내 세번째 시작 행
    location_Col_1 = "B" # 고정 열
    location_Col_2 = "U" # 번호확인 함수 위치
    location_Col_3 = "O" # 설명확인 함수 위치
    imageCycle = 1   # 각 동의 폴더에서 넣을 사진의 순서
    contentCycle = 1 # 조사표 내용이 삽입되는 순서
    fileList = None
    fileListJpg = None


    def setExcelFilePath(self, excelFilePath):
        self.excelFilePath = excelFilePath
    def getExcelFilePath(self):
        return self.excelFilePath

    def setSheetsNameList(self, sheetName):
        self.sheetsNameList.append(sheetName)
    def getSheetsNameList(self):
        return self.sheetsNameList

    def setImgFilePathList(self, imgFilePath):
        self.imgFilePathList.append(imgFilePath)
    def getImgFilePathList(self):
        return self.imgFilePathList

    def setImgFileList(self):
        for index, filePath in enumerate(self.getImgFilePathList()):
            tmpList = []
            tmpList = os.listdir(filePath)
            self.imgFileList.append([])
            for tmp in tmpList:
                if (tmp.endswith(".jpg") or tmp.endswith(".JPG")):
                    self.imgFileList[index].append(tmp)
              
    def getImgFileList(self):
        return self.imgFileList

    def setImageNum(self):
        for size in self.getImgFileList():
            self.imageNum.append(len(size))

    def getImageNum(self):
        return self.imageNum
    
    def setPageNum(self):
        for size in self.getImageNum():
            self.pageNum.append(size // 3) # 사진이 무조건 다 들어가는 페이지수
    def getPageNum(self):
        return self.pageNum
    
    def setLastPageNum(self):
        for size in self.getImageNum():
            self.lastPageNum.append(3 - (size % 3))
 # 마지막 페이지에서 모자란 사진 숫자
    def getLastPageNum(self):
        return self.lastPageNum
    

    def setWorkBook(self):

        self.wb = self.excel.Workbooks.Open(self.getExcelFilePath())
    def getWorkBook(self):
        return self.wb

    def setWorkSheets(self):
        for index in self.getSheetsNameList():
            if self.wb.Sheets(index):
                self.ws.append(self.wb.Sheets(index))
        

    def getWorkSheets(self):
        return self.ws

    def setWorkSheetsName(self):
        # self.ws = self.wb.Sheets(self.sheetsName)
        self.wsName = [sheet.Name for sheet in self.wb.Sheets]
        
    def getWorkSheetsName(self):
        return self.wsName

# 이미지를 복사해주는 함수    
    def duplicateImage(self):
        for index, size in enumerate(self.getLastPageNum()):
            if ((size < 3) and (size > 0)):
                self.pageNum[index] += 1
                for k in range(1, self.lastPageNum[index]+1, 1):
                    shutil.copyfile(self.getImgFilePathList()[index] +  "\\" + str(self.imageNum[index]) + ".jpg",self.getImgFilePathList()[index] +  "\\" + str(self.imageNum[index]+1) + ".jpg")
                    self.imgFileList[index].append(str(self.imageNum[index]+1) + ".jpg")
           
    
    # def duplicateImage(self):
    #     try:
    #         for index, size in enumerate(self.getLastPageNum()):
    #             if ((size < 3) and (size > 0)):
    #                 self.pageNum[index] += 1
    #                 for k in range(1, self.lastPageNum[index]+1, 1):
    #                     shutil.copyfile(self.getImgFilePathList()[index] +  "\\" + str(self.imageNum[index]) + ".jpg",self.getImgFilePathList()[index] +  "\\" + str(self.imageNum[index]+1) + ".jpg")
    #                     self.imgFileList[index].append(str(self.imageNum[index]+1) + ".jpg")
    #                     self.imageNum[index] += 1
    #         return None
    #     except Exception as e:
    #         errorMessage = "에러 발생", str(self.getWorkSheets()[index]) + "에 들어갈 " + str(self.imageNum[index]) + ".jpg 사진이 없습니다."
    #         return errorMessage

    # 화살표를 복사해주는 함수
    def inputArrow(self, location_Row, index):
        self.ws[index].Range("Z1:AF4").Copy(self.ws[index].Range("U"+str(location_Row)))
    

    # 결함의 위치를 삽입해주는 함수
    def inputLocation(self, location_Row, index):
        self.ws[index].Cells(int(location_Row)+1, 17).Value = (str(self.ws[index].Cells(int(self.contentCycle)+8, 33)))+"\n"+(str(self.ws[index].Cells(int(self.contentCycle)+8, 35)))
        return ()


    # 조사표 내용을 조사사진에 넣어주는 함수
    def inputContents(self,location_Row, index):
        #41,43,45,47,49    25,38
        self.ws[index].Cells(int(location_Row)+8, 15).Value = self.ws[index].Cells(int(self.contentCycle)+9, 41)
        self.ws[index].Cells(int(location_Row)+8, 17).Value = self.ws[index].Cells(int(self.contentCycle)+9, 43)
        self.ws[index].Cells(int(location_Row)+8, 19).Value = self.ws[index].Cells(int(self.contentCycle)+9, 45)
        self.ws[index].Cells(int(location_Row)+8, 21).Value = self.ws[index].Cells(int(self.contentCycle)+9, 47)
        self.ws[index].Cells(int(location_Row)+8, 23).Value = self.ws[index].Cells(int(self.contentCycle)+9, 49)
        self.ws[index].Cells(int(location_Row)+8, 25).Value = self.ws[index].Cells(int(self.contentCycle)+9, 38)
        self.contentCycle = self.contentCycle + 1
        return(self.contentCycle)


    # 설명번호와 사진번호가 맞는지 확인해주는 엑셀수식을 삽입해주는 함수
    def checkImageNum(self, location_Row, index):
        self.ws[index].Cells(int(location_Row)+1, 21).Value = "=IF(MID(O"+str(int(location_Row)+1)+",SEARCH(\".\",O"+str(int(location_Row)+1)+",1),3)=MID(W"+str(int(location_Row)+8)+",SEARCH(\".\",W"+str(int(location_Row)+8)+",1),3),\"\",\"번호확인\")" 
        self.ws[index].Cells(int(location_Row)+1, 21).Font.Color = -16776961
        return()


    # 설명부분 내용을 합성해주는 엑셀수식을 삽입해주는 함수
    def combineExplanation(self, location_Row, location_Col, index):
        if(self.ws[index].Cells(int(location_Row)+8, 25).Value == "균열"):
            self.ws[index].Cells(int(location_Row)+4, 15).Value = "="+ location_Col + str(int(location_Row) + 8) + "&"+"\"균열\"" 
        else:
            self.ws[index].Cells(int(location_Row)+4, 15).Value = self.ws[index].Cells(int(location_Row)+8, 25)
        return()


    # 이전 행(Row)을 받고 다음 사진이 삽입될 행(Row)의 넘버를 반환해주는 함수
    def nextLocation(self, location_Row):
        location_Row = location_Row + 32
        return(location_Row)


    # 행, 열, 사진경로, 폴더이름, 사진 번호를 받고 사진을 해당위치에 삽입 및 다음 사진 넘버를 반환 해주는 함수
    def inputImage(self, location_Row, location_Col, index):
        location = location_Col + str(location_Row)
        rng = self.ws[index].Range(location) 
        ImagePath = self.getImgFilePathList()[index] + "/" + str(self.imageCycle) + ".jpg"
        ImagePath = ImagePath.replace("/","\\")
        self.ws[index].Shapes.AddPicture(ImagePath, False,True, rng.Left, rng.Top, 247.68, 184.28)
        self.imageCycle = self.imageCycle + 1
        return(self.imageCycle) 