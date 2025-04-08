namespace XlsToXlsxConverter
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.mainLayout = new System.Windows.Forms.TableLayoutPanel();
            this.lblFolder = new System.Windows.Forms.Label();
            this.cmbFolderPath = new System.Windows.Forms.ComboBox();
            this.btnSelectFolder = new System.Windows.Forms.Button();
            this.chkDeleteOriginal = new System.Windows.Forms.CheckBox();
            this.chkOverwrite = new System.Windows.Forms.CheckBox();
            this.buttonPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.btnStartConversion = new System.Windows.Forms.Button();
            this.btnCancelConversion = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lstResults = new System.Windows.Forms.ListView();
            this.colFileName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colStatus = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colSourcePath = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colDestinationPath = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colErrorInfo = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.mainLayout.SuspendLayout();
            this.buttonPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // mainLayout
            // 
            this.mainLayout.ColumnCount = 3;
            this.mainLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.mainLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 65F));
            this.mainLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15F));
            this.mainLayout.Controls.Add(this.lblFolder, 0, 0);
            this.mainLayout.Controls.Add(this.cmbFolderPath, 1, 0);
            this.mainLayout.Controls.Add(this.btnSelectFolder, 2, 0);
            this.mainLayout.Controls.Add(this.chkDeleteOriginal, 1, 1);
            this.mainLayout.Controls.Add(this.chkOverwrite, 1, 2);
            this.mainLayout.Controls.Add(this.buttonPanel, 1, 3);
            this.mainLayout.Controls.Add(this.lblStatus, 1, 4);
            this.mainLayout.Controls.Add(this.progressBar, 1, 5);
            this.mainLayout.Controls.Add(this.lstResults, 0, 6);
            this.mainLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainLayout.Location = new System.Drawing.Point(0, 0);
            this.mainLayout.Name = "mainLayout";
            this.mainLayout.Padding = new System.Windows.Forms.Padding(10);
            this.mainLayout.RowCount = 7;
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.mainLayout.Size = new System.Drawing.Size(784, 561);
            this.mainLayout.TabIndex = 0;
            // 
            // lblFolder
            // 
            this.lblFolder.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFolder.Location = new System.Drawing.Point(13, 10);
            this.lblFolder.Name = "lblFolder";
            this.lblFolder.Size = new System.Drawing.Size(146, 30);
            this.lblFolder.TabIndex = 0;
            this.lblFolder.Text = "フォルダパス:";
            this.lblFolder.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cmbFolderPath
            // 
            this.cmbFolderPath.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cmbFolderPath.FormattingEnabled = true;
            this.cmbFolderPath.Location = new System.Drawing.Point(165, 13);
            this.cmbFolderPath.Name = "cmbFolderPath";
            this.cmbFolderPath.Size = new System.Drawing.Size(491, 21);
            this.cmbFolderPath.TabIndex = 1;
            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnSelectFolder.Location = new System.Drawing.Point(662, 13);
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Size = new System.Drawing.Size(109, 24);
            this.btnSelectFolder.TabIndex = 2;
            this.btnSelectFolder.Text = "選択...";
            this.btnSelectFolder.UseVisualStyleBackColor = true;
            this.btnSelectFolder.Click += new System.EventHandler(this.btnSelectFolder_Click);
            // 
            // chkDeleteOriginal
            // 
            this.chkDeleteOriginal.AutoSize = true;
            this.chkDeleteOriginal.Location = new System.Drawing.Point(165, 43);
            this.chkDeleteOriginal.Name = "chkDeleteOriginal";
            this.chkDeleteOriginal.Size = new System.Drawing.Size(180, 17);
            this.chkDeleteOriginal.TabIndex = 3;
            this.chkDeleteOriginal.Text = "変換後に元のXLSファイルを削除する";
            this.chkDeleteOriginal.UseVisualStyleBackColor = true;
            // 
            // chkOverwrite
            // 
            this.chkOverwrite.AutoSize = true;
            this.chkOverwrite.Checked = true;
            this.chkOverwrite.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkOverwrite.Location = new System.Drawing.Point(165, 73);
            this.chkOverwrite.Name = "chkOverwrite";
            this.chkOverwrite.Size = new System.Drawing.Size(204, 17);
            this.chkOverwrite.TabIndex = 4;
            this.chkOverwrite.Text = "既存のXLSXファイルが存在する場合は上書き";
            this.chkOverwrite.UseVisualStyleBackColor = true;
            // 
            // buttonPanel
            // 
            this.buttonPanel.Controls.Add(this.btnStartConversion);
            this.buttonPanel.Controls.Add(this.btnCancelConversion);
            this.buttonPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.buttonPanel.FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight;
            this.buttonPanel.Location = new System.Drawing.Point(165, 103);
            this.buttonPanel.Name = "buttonPanel";
            this.buttonPanel.Size = new System.Drawing.Size(491, 34);
            this.buttonPanel.TabIndex = 5;
            // 
            // btnStartConversion
            // 
            this.btnStartConversion.Location = new System.Drawing.Point(3, 3);
            this.btnStartConversion.Name = "btnStartConversion";
            this.btnStartConversion.Size = new System.Drawing.Size(120, 30);
            this.btnStartConversion.TabIndex = 0;
            this.btnStartConversion.Text = "変換開始";
            this.btnStartConversion.UseVisualStyleBackColor = true;
            this.btnStartConversion.Click += new System.EventHandler(this.btnStartConversion_Click);
            // 
            // btnCancelConversion
            // 
            this.btnCancelConversion.Enabled = false;
            this.btnCancelConversion.Location = new System.Drawing.Point(129, 3);
            this.btnCancelConversion.Name = "btnCancelConversion";
            this.btnCancelConversion.Size = new System.Drawing.Size(120, 30);
            this.btnCancelConversion.TabIndex = 1;
            this.btnCancelConversion.Text = "変換中止";
            this.btnCancelConversion.UseVisualStyleBackColor = true;
            this.btnCancelConversion.Click += new System.EventHandler(this.btnCancelConversion_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblStatus.Location = new System.Drawing.Point(165, 140);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(491, 30);
            this.lblStatus.TabIndex = 6;
            this.lblStatus.Text = "準備完了";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // progressBar
            // 
            this.progressBar.Dock = System.Windows.Forms.DockStyle.Fill;
            this.progressBar.Location = new System.Drawing.Point(165, 173);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(491, 24);
            this.progressBar.TabIndex = 7;
            // 
            // lstResults
            // 
            this.lstResults.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colFileName,
            this.colStatus,
            this.colSourcePath,
            this.colDestinationPath,
            this.colErrorInfo});
            this.mainLayout.SetColumnSpan(this.lstResults, 3);
            this.lstResults.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lstResults.FullRowSelect = true;
            this.lstResults.GridLines = true;
            this.lstResults.HideSelection = false;
            this.lstResults.Location = new System.Drawing.Point(13, 203);
            this.lstResults.Name = "lstResults";
            this.lstResults.Size = new System.Drawing.Size(758, 345);
            this.lstResults.TabIndex = 8;
            this.lstResults.UseCompatibleStateImageBehavior = false;
            this.lstResults.View = System.Windows.Forms.View.Details;
            this.lstResults.DoubleClick += new System.EventHandler(this.lstResults_DoubleClick);
            // 
            // colFileName
            // 
            this.colFileName.Text = "ファイル名";
            this.colFileName.Width = 200;
            // 
            // colStatus
            // 
            this.colStatus.Text = "ステータス";
            this.colStatus.Width = 80;
            // 
            // colSourcePath
            // 
            this.colSourcePath.Text = "元ファイルパス";
            this.colSourcePath.Width = 200;
            // 
            // colDestinationPath
            // 
            this.colDestinationPath.Text = "変換先ファイルパス";
            this.colDestinationPath.Width = 200;
            // 
            // colErrorInfo
            // 
            this.colErrorInfo.Text = "エラー情報";
            this.colErrorInfo.Width = 200;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 561);
            this.Controls.Add(this.mainLayout);
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "XLS to XLSX コンバーター";
            this.mainLayout.ResumeLayout(false);
            this.mainLayout.PerformLayout();
            this.buttonPanel.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel mainLayout;
        private System.Windows.Forms.Label lblFolder;
        private System.Windows.Forms.ComboBox cmbFolderPath;
        private System.Windows.Forms.Button btnSelectFolder;
        private System.Windows.Forms.CheckBox chkDeleteOriginal;
        private System.Windows.Forms.CheckBox chkOverwrite;
        private System.Windows.Forms.FlowLayoutPanel buttonPanel;
        private System.Windows.Forms.Button btnStartConversion;
        private System.Windows.Forms.Button btnCancelConversion;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.ListView lstResults;
        private System.Windows.Forms.ColumnHeader colFileName;
        private System.Windows.Forms.ColumnHeader colStatus;
        private System.Windows.Forms.ColumnHeader colSourcePath;
        private System.Windows.Forms.ColumnHeader colDestinationPath;
        private System.Windows.Forms.ColumnHeader colErrorInfo;
    }
}