using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;



namespace ExcelAutoInput
{
    public partial class Form1 : Form
    {

        private ExcelDocumentInfo excelFileInfo = null;
        private DirectoryInfo di = null;

        


        public Form1()
        {
            InitializeComponent();
            btnSelectSheet.Enabled = false;
            btnSelectRootFolder.Enabled = false;
            btnSelectImgFolder.Enabled = false;
            btnInputImage.Enabled = false;
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            excelFileInfo = new ExcelDocumentInfo();
            openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            excelFileInfo.SetFileName(openFileDialog1.FileName);
            excelFileInfo.SetWorkSheetList(excelFileInfo.OpenWorkbook());
            for (int i = 0; i < excelFileInfo.GetWorkSheetList().Count; i++)
            {
                checkedListBox1.Items.Add(excelFileInfo.GetWorkSheetList()[i].Name);
            }
            btnOpenFile.Enabled = false;
            btnSelectSheet.Enabled = true;

        }
        private void btnSelectSheet_Click(object sender, EventArgs e)
        {
            foreach (string checkedSheet in checkedListBox1.CheckedItems)
            {
                foreach(Excel.Worksheet workSheet in excelFileInfo.GetWorkSheetList())
                if(workSheet.Name == checkedSheet)
                    {
                        excelFileInfo.SetSelectedSheetList(workSheet);
                    }
            }
            checkedListBox1.Enabled = false;
            btnSelectSheet.Enabled = false;
            btnSelectRootFolder.Enabled = true;

        }


        private void btnSelectRootFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.ShowDialog();
            di = new DirectoryInfo(@folderBrowserDialog.SelectedPath);
            foreach (DirectoryInfo directory in di.EnumerateDirectories())
            {

                checkedListBox2.Items.Add(directory.Name);
            }
            btnSelectRootFolder.Enabled = false;
            btnSelectImgFolder.Enabled = true;
        }

        private void btnInputImage_Click(object sender, EventArgs e)
        {
            // 이미지의 갯수가 3의 배수가 아니면 복사하는 함수
            for (int i = 0; i < checkedListBox2.CheckedItems.Count; i++)
            {
                excelFileInfo.DuplicateImg(excelFileInfo.GetSelectedImgFolderPathList()[i]);
            }
            excelFileInfo.InputImg();
            // 화살표 넣어주는 함수
            // 조사표 내용 넣어주는 함수
            // 사진 번호 수식 넣어주는 함수
            // 결함 위치 내용 합쳐주는 함수
            // 사진 넣어주는 함수
            btnInputImage.Enabled = false;
        }



        private void btnExit_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void btnSelectImgFolder_Click(object sender, EventArgs e)
        {
            // 선택한 이미지 폴더 정보를 넘겨줘야함
            foreach (string folderPath in checkedListBox2.CheckedItems)
            {
                excelFileInfo.SetSelectedImgFolderPathList(di.FullName + "\\" + folderPath);
            }
            checkedListBox2.Enabled = false;
            btnSelectImgFolder.Enabled = false;
            btnInputImage.Enabled = true;
        }
    }
}
