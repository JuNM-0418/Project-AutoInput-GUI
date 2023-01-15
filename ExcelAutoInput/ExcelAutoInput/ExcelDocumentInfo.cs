using System.Collections.Generic;
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
        private List<string> SelectedSheetList = null;
        private List<string> ImgFolderPathList = null;

        // 이미지의 갯수가 3의 배수가 아니면 복사하는 함수
        // 조사표 내용 넣어주는 함수
        // 사진 번호 수식 넣어주는 함수
        // 결함 위치 내용 합쳐주는 함수
        // 사진 넣어주는 함수

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


        public void SetSelectedSheetList(string selectedSheet)
        {
            if (SelectedSheetList == null)
            {
                SelectedSheetList = new List<string>();
            }
            this.SelectedSheetList.Add(selectedSheet);

        }
        public List<string> GetSelectedSheetList()
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
