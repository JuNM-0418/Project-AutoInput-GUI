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
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.btnSelectSheet = new System.Windows.Forms.Button();
            this.btnSelectRootFolder = new System.Windows.Forms.Button();
            this.btnInputImage = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnExit = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.checkedListBox2 = new System.Windows.Forms.CheckedListBox();
            this.btnSelectImgFolder = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Location = new System.Drawing.Point(12, 12);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(208, 32);
            this.btnOpenFile.TabIndex = 0;
            this.btnOpenFile.Text = "파일 열기";
            this.btnOpenFile.UseVisualStyleBackColor = true;
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(12, 50);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(208, 228);
            this.checkedListBox1.TabIndex = 2;
            // 
            // btnSelectSheet
            // 
            this.btnSelectSheet.Location = new System.Drawing.Point(12, 284);
            this.btnSelectSheet.Name = "btnSelectSheet";
            this.btnSelectSheet.Size = new System.Drawing.Size(208, 32);
            this.btnSelectSheet.TabIndex = 0;
            this.btnSelectSheet.Text = "시트 선택 완료";
            this.btnSelectSheet.UseVisualStyleBackColor = true;
            this.btnSelectSheet.Click += new System.EventHandler(this.btnSelectSheet_Click);
            // 
            // btnSelectRootFolder
            // 
            this.btnSelectRootFolder.Location = new System.Drawing.Point(226, 12);
            this.btnSelectRootFolder.Name = "btnSelectRootFolder";
            this.btnSelectRootFolder.Size = new System.Drawing.Size(208, 32);
            this.btnSelectRootFolder.TabIndex = 3;
            this.btnSelectRootFolder.Text = "사진 폴더 선택";
            this.btnSelectRootFolder.UseVisualStyleBackColor = true;
            this.btnSelectRootFolder.Click += new System.EventHandler(this.btnSelectRootFolder_Click);
            // 
            // btnInputImage
            // 
            this.btnInputImage.Location = new System.Drawing.Point(12, 322);
            this.btnInputImage.Name = "btnInputImage";
            this.btnInputImage.Size = new System.Drawing.Size(422, 32);
            this.btnInputImage.TabIndex = 4;
            this.btnInputImage.Text = "작업 시작";
            this.btnInputImage.UseVisualStyleBackColor = true;
            this.btnInputImage.Click += new System.EventHandler(this.btnInputImage_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(12, 398);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(422, 32);
            this.btnExit.TabIndex = 5;
            this.btnExit.Text = "종료";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 360);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(422, 32);
            this.progressBar1.TabIndex = 6;
            // 
            // checkedListBox2
            // 
            this.checkedListBox2.FormattingEnabled = true;
            this.checkedListBox2.Location = new System.Drawing.Point(226, 50);
            this.checkedListBox2.Name = "checkedListBox2";
            this.checkedListBox2.Size = new System.Drawing.Size(208, 228);
            this.checkedListBox2.TabIndex = 7;
            // 
            // btnSelectImgFolder
            // 
            this.btnSelectImgFolder.Location = new System.Drawing.Point(226, 284);
            this.btnSelectImgFolder.Name = "btnSelectImgFolder";
            this.btnSelectImgFolder.Size = new System.Drawing.Size(208, 32);
            this.btnSelectImgFolder.TabIndex = 8;
            this.btnSelectImgFolder.Text = "사진 폴더 선택 완료";
            this.btnSelectImgFolder.UseVisualStyleBackColor = true;
            this.btnSelectImgFolder.Click += new System.EventHandler(this.btnSelectImgFolder_Click);
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.HelpRequest += new System.EventHandler(this.folderBrowserDialog1_HelpRequest);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(444, 506);
            this.Controls.Add(this.btnSelectImgFolder);
            this.Controls.Add(this.checkedListBox2);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnInputImage);
            this.Controls.Add(this.btnSelectRootFolder);
            this.Controls.Add(this.checkedListBox1);
            this.Controls.Add(this.btnSelectSheet);
            this.Controls.Add(this.btnOpenFile);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.Button btnSelectSheet;
        private System.Windows.Forms.Button btnSelectRootFolder;
        private System.Windows.Forms.Button btnInputImage;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.CheckedListBox checkedListBox2;
        private System.Windows.Forms.Button btnSelectImgFolder;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }
}

