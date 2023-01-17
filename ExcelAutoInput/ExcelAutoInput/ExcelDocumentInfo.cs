using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAutoInput
{
    internal class ExcelDocumentInfo
    {
        private Excel.Application application = null;
        private string FileName = null;
        private Excel.Workbook WorkBook = null;
        private Excel.Worksheet WorkSheet = null;
        private List<Excel.Worksheet> WorkSheetList = null;
        private List<Excel.Worksheet> SelectedSheetList = null;
        private List<string> ImgFolderPathList = null;

        private int LocationRowTop = 3; // 페이지 내 첫번째 시작 행
        private int LocationRowMid = 13; // 페이지 내 두번쨰 시작 행
        private int LocationRowBottom = 23; // 페이지 내 세번째 시작 행
        private int LocationColImg = 2; // 사진 삽입 고정 열
        private string LocationColCheckNum = "U"; // 번호확인 함수 위치
        private string LocationColExplainFunction = "O"; // 설명확인 함수 위치
        private int ImgCycle = 1;   // 각 동의 폴더에서 넣을 사진의 순서
        private int SurveyDataCycle = 1; // 조사표 내용이 삽입되는 순서

        // 다음 행의 위치를 반환하는 함수
        public int NextLocationRow(int locationRow)
        {
            locationRow += 32;
            return locationRow;
        }

        // 사진 넣어주는 함수
        public void InputImg()
        {
            this.LocationRowTop= 3;
            this.LocationColImg = 2;
            this.ImgCycle = 1;
            /**
            foreach(Excel.Worksheet workSheet in this.GetSelectedSheetList())
            {

            }
            */
            Excel.Range range = this.GetSelectedSheetList()[0].Cells[this.LocationRowTop, this.LocationColImg];
            string imgPath = this.GetSelectedImgFolderPathList()[0] + "\\" + this.ImgCycle.ToString() + ".jpg";
            this.GetSelectedSheetList()[0].Shapes.AddPicture(@imgPath, 0, (Microsoft.Office.Core.MsoTriState)1, range.Left, range.Top, (float)246.68, (float)184.28);
        }

        // 화살표를 넣어주는 함수
        public void InputArrow()
        {

        }
       

        // 결함 위치 내용 합쳐주는 수식을 넣어주는 함수
        public void InputCombineExcelFuncion()
        {

        }

        // 사진 번호 수식 넣어주는 함수
        public void InputImgNum()
        {

        }

        // 조사표 내용 넣어주는 함수
        public void InputSurveyData()
        {

        }

        // 이미지의 갯수가 3의 배수가 아니면 복사하는 함수
        public void DuplicateImg(string folderPath)
        {
            DirectoryInfo di = new DirectoryInfo(@folderPath);
            IEnumerable<FileInfo> imgList = di.EnumerateFiles("*.jpg", SearchOption.AllDirectories);
            if (imgList.Count() % 3 != 0)
            {
                int addImgNum = 3 - (imgList.Count() % 3);
                for (int i = 1; i <= addImgNum; i++)
                {
                    File.Copy(imgList.First().FullName, imgList.First().Directory + "\\" + (imgList.Count() + 1) + ".jpg");
                }
            }
        }


        public void SetFileName(string fileName)
        {
            this.FileName = fileName;
        }
        public string GetFileName()
        {
            return this.FileName;
        }


        public Excel.Workbook OpenWorkbook()
        {
            this.WorkBook = application.Workbooks.Open(GetFileName());
            return this.WorkBook;
        }


        public void SetWorkSheetList(Excel.Workbook workBook)
        {
            this.WorkSheetList = new List<Excel.Worksheet>();
            for (int i = 1; i <= this.WorkBook.Sheets.Count; i++)
            {
                this.WorkSheetList.Add(workBook.Sheets.get_Item(i));
            }
        }
        public List<Excel.Worksheet> GetWorkSheetList()
        {
            return this.WorkSheetList;
        }


        public void SetSelectedSheetList(Excel.Worksheet selectedSheet)
        {
            if (SelectedSheetList == null)
            {
                SelectedSheetList = new List<Excel.Worksheet>();
            }
            this.SelectedSheetList.Add(selectedSheet);

        }
        public List<Excel.Worksheet> GetSelectedSheetList()
        {
            return this.SelectedSheetList;
        }


        public void SetSelectedImgFolderPathList(string folderPath)
        {
            if (ImgFolderPathList == null)
            {
                ImgFolderPathList = new List<string>();
            }
            this.ImgFolderPathList.Add(folderPath);

        }
        public List<string> GetSelectedImgFolderPathList()
        {
            return this.ImgFolderPathList;
        }


        public ExcelDocumentInfo()
        {
            this.application = new Excel.Application();
            this.application.Visible = true;

        }
        ~ExcelDocumentInfo()
        {
            this.WorkBook.Close();
            this.application.Quit();
            Marshal.ReleaseComObject(this.WorkBook);
            Marshal.ReleaseComObject(this.application);
        }

    }
}
