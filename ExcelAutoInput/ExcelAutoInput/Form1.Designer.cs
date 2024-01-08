namespace ExcelAutoInput
{
    partial class Form1
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.sheetListBox = new System.Windows.Forms.CheckedListBox();
            this.btnSelectSheet = new System.Windows.Forms.Button();
            this.btnSelectRootFolder = new System.Windows.Forms.Button();
            this.btnInputImage = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnExit = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.imgPathListBox = new System.Windows.Forms.CheckedListBox();
            this.btnSelectImgFolder = new System.Windows.Forms.Button();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.surveySheetListBox = new System.Windows.Forms.CheckedListBox();
            this.btnSelectSurveySheet = new System.Windows.Forms.Button();
            this.btnSelectSurveySheetDone = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Location = new System.Drawing.Point(12, 12);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(420, 32);
            this.btnOpenFile.TabIndex = 0;
            this.btnOpenFile.Text = "파일 열기";
            this.btnOpenFile.UseVisualStyleBackColor = true;
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // sheetListBox
            // 
            this.sheetListBox.FormattingEnabled = true;
            this.sheetListBox.Location = new System.Drawing.Point(12, 50);
            this.sheetListBox.Name = "sheetListBox";
            this.sheetListBox.Size = new System.Drawing.Size(420, 84);
            this.sheetListBox.TabIndex = 2;
            // 
            // btnSelectSheet
            // 
            this.btnSelectSheet.Location = new System.Drawing.Point(12, 140);
            this.btnSelectSheet.Name = "btnSelectSheet";
            this.btnSelectSheet.Size = new System.Drawing.Size(420, 32);
            this.btnSelectSheet.TabIndex = 0;
            this.btnSelectSheet.Text = "시트 선택 완료";
            this.btnSelectSheet.UseVisualStyleBackColor = true;
            this.btnSelectSheet.Click += new System.EventHandler(this.btnSelectSheet_Click);
            // 
            // btnSelectRootFolder
            // 
            this.btnSelectRootFolder.Location = new System.Drawing.Point(12, 178);
            this.btnSelectRootFolder.Name = "btnSelectRootFolder";
            this.btnSelectRootFolder.Size = new System.Drawing.Size(420, 32);
            this.btnSelectRootFolder.TabIndex = 3;
            this.btnSelectRootFolder.Text = "사진 폴더 선택";
            this.btnSelectRootFolder.UseVisualStyleBackColor = true;
            this.btnSelectRootFolder.Click += new System.EventHandler(this.btnSelectRootFolder_Click);
            // 
            // btnInputImage
            // 
            this.btnInputImage.Location = new System.Drawing.Point(12, 478);
            this.btnInputImage.Name = "btnInputImage";
            this.btnInputImage.Size = new System.Drawing.Size(422, 32);
            this.btnInputImage.TabIndex = 4;
            this.btnInputImage.Text = "작업 시작";
            this.btnInputImage.UseVisualStyleBackColor = true;
            this.btnInputImage.Click += new System.EventHandler(this.btnInputImage_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog";
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(12, 554);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(422, 32);
            this.btnExit.TabIndex = 5;
            this.btnExit.Text = "종료";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 516);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(422, 32);
            this.progressBar.TabIndex = 6;
            // 
            // imgPathListBox
            // 
            this.imgPathListBox.FormattingEnabled = true;
            this.imgPathListBox.Location = new System.Drawing.Point(12, 216);
            this.imgPathListBox.Name = "imgPathListBox";
            this.imgPathListBox.Size = new System.Drawing.Size(420, 68);
            this.imgPathListBox.TabIndex = 7;
            // 
            // btnSelectImgFolder
            // 
            this.btnSelectImgFolder.Location = new System.Drawing.Point(12, 290);
            this.btnSelectImgFolder.Name = "btnSelectImgFolder";
            this.btnSelectImgFolder.Size = new System.Drawing.Size(420, 32);
            this.btnSelectImgFolder.TabIndex = 8;
            this.btnSelectImgFolder.Text = "사진 폴더 선택 완료";
            this.btnSelectImgFolder.UseVisualStyleBackColor = true;
            this.btnSelectImgFolder.Click += new System.EventHandler(this.btnSelectImgFolder_Click);
            // 
            // surveySheetListBox
            // 
            this.surveySheetListBox.FormattingEnabled = true;
            this.surveySheetListBox.Location = new System.Drawing.Point(12, 366);
            this.surveySheetListBox.Name = "surveySheetListBox";
            this.surveySheetListBox.Size = new System.Drawing.Size(420, 68);
            this.surveySheetListBox.TabIndex = 7;
            // 
            // btnSelectSurveySheet
            // 
            this.btnSelectSurveySheet.Location = new System.Drawing.Point(10, 328);
            this.btnSelectSurveySheet.Name = "btnSelectSurveySheet";
            this.btnSelectSurveySheet.Size = new System.Drawing.Size(422, 32);
            this.btnSelectSurveySheet.TabIndex = 9;
            this.btnSelectSurveySheet.Text = "조사표 시트 선택";
            this.btnSelectSurveySheet.UseVisualStyleBackColor = true;
            this.btnSelectSurveySheet.Click += new System.EventHandler(this.btnSelectSurveySheet_Click);
            // 
            // btnSelectSurveySheetDone
            // 
            this.btnSelectSurveySheetDone.Location = new System.Drawing.Point(10, 440);
            this.btnSelectSurveySheetDone.Name = "btnSelectSurveySheetDone";
            this.btnSelectSurveySheetDone.Size = new System.Drawing.Size(422, 32);
            this.btnSelectSurveySheetDone.TabIndex = 9;
            this.btnSelectSurveySheetDone.Text = "조사표 시트 선택 완료";
            this.btnSelectSurveySheetDone.UseVisualStyleBackColor = true;
            this.btnSelectSurveySheetDone.Click += new System.EventHandler(this.btnSelectSurveySheetDone_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(444, 651);
            this.Controls.Add(this.btnSelectSurveySheetDone);
            this.Controls.Add(this.btnSelectSurveySheet);
            this.Controls.Add(this.btnSelectImgFolder);
            this.Controls.Add(this.surveySheetListBox);
            this.Controls.Add(this.imgPathListBox);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnInputImage);
            this.Controls.Add(this.btnSelectRootFolder);
            this.Controls.Add(this.sheetListBox);
            this.Controls.Add(this.btnSelectSheet);
            this.Controls.Add(this.btnOpenFile);
            this.Name = "Form1";
            this.Text = "안전점검";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.CheckedListBox sheetListBox;
        private System.Windows.Forms.Button btnSelectSheet;
        private System.Windows.Forms.Button btnSelectRootFolder;
        private System.Windows.Forms.Button btnInputImage;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.CheckedListBox imgPathListBox;
        private System.Windows.Forms.Button btnSelectImgFolder;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.CheckedListBox surveySheetListBox;
        private System.Windows.Forms.Button btnSelectSurveySheet;
        private System.Windows.Forms.Button btnSelectSurveySheetDone;
    }
}

