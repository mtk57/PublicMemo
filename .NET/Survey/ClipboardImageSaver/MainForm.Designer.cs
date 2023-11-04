namespace ClipboardImageSaver
{
    partial class MainForm
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxSaveDirPath = new System.Windows.Forms.TextBox();
            this.buttonSaveDirRef = new System.Windows.Forms.Button();
            this.textBoxSaveFileName = new System.Windows.Forms.TextBox();
            this.numericUpDownStartNum = new System.Windows.Forms.NumericUpDown();
            this.buttonResume = new System.Windows.Forms.Button();
            this.textBoxLog = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownStartNum)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(42, 194);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(348, 24);
            this.label1.TabIndex = 0;
            this.label1.Text = "保存ファイル名 (拡張子は含めない)";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(42, 323);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(106, 24);
            this.label2.TabIndex = 1;
            this.label2.Text = "開始番号";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(42, 63);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(234, 24);
            this.label3.TabIndex = 2;
            this.label3.Text = "保存フォルダ (絶対パス)";
            // 
            // textBoxSaveDirPath
            // 
            this.textBoxSaveDirPath.Location = new System.Drawing.Point(46, 99);
            this.textBoxSaveDirPath.Name = "textBoxSaveDirPath";
            this.textBoxSaveDirPath.Size = new System.Drawing.Size(776, 31);
            this.textBoxSaveDirPath.TabIndex = 3;
            // 
            // buttonSaveDirRef
            // 
            this.buttonSaveDirRef.Location = new System.Drawing.Point(838, 99);
            this.buttonSaveDirRef.Name = "buttonSaveDirRef";
            this.buttonSaveDirRef.Size = new System.Drawing.Size(81, 35);
            this.buttonSaveDirRef.TabIndex = 4;
            this.buttonSaveDirRef.Text = "参照";
            this.buttonSaveDirRef.UseVisualStyleBackColor = true;
            this.buttonSaveDirRef.Click += new System.EventHandler(this.ButtonSaveDirRef_Click);
            // 
            // textBoxSaveFileName
            // 
            this.textBoxSaveFileName.Location = new System.Drawing.Point(46, 230);
            this.textBoxSaveFileName.Name = "textBoxSaveFileName";
            this.textBoxSaveFileName.Size = new System.Drawing.Size(776, 31);
            this.textBoxSaveFileName.TabIndex = 5;
            // 
            // numericUpDownStartNum
            // 
            this.numericUpDownStartNum.Location = new System.Drawing.Point(46, 359);
            this.numericUpDownStartNum.Maximum = new decimal(new int[] {
            1215752191,
            23,
            0,
            0});
            this.numericUpDownStartNum.Name = "numericUpDownStartNum";
            this.numericUpDownStartNum.Size = new System.Drawing.Size(173, 31);
            this.numericUpDownStartNum.TabIndex = 6;
            this.numericUpDownStartNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // buttonResume
            // 
            this.buttonResume.Location = new System.Drawing.Point(280, 426);
            this.buttonResume.Name = "buttonResume";
            this.buttonResume.Size = new System.Drawing.Size(402, 92);
            this.buttonResume.TabIndex = 7;
            this.buttonResume.Text = "クリップボード監視を再開する";
            this.buttonResume.UseVisualStyleBackColor = true;
            this.buttonResume.Click += new System.EventHandler(this.ButtonResume_Click);
            // 
            // textBoxLog
            // 
            this.textBoxLog.Location = new System.Drawing.Point(37, 548);
            this.textBoxLog.Multiline = true;
            this.textBoxLog.Name = "textBoxLog";
            this.textBoxLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxLog.Size = new System.Drawing.Size(893, 186);
            this.textBoxLog.TabIndex = 8;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(964, 764);
            this.Controls.Add(this.textBoxLog);
            this.Controls.Add(this.buttonResume);
            this.Controls.Add(this.numericUpDownStartNum);
            this.Controls.Add(this.textBoxSaveFileName);
            this.Controls.Add(this.buttonSaveDirRef);
            this.Controls.Add(this.textBoxSaveDirPath);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.Text = "Clipboard Image Saver";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Shown += new System.EventHandler(this.MainForm_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownStartNum)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxSaveDirPath;
        private System.Windows.Forms.Button buttonSaveDirRef;
        private System.Windows.Forms.TextBox textBoxSaveFileName;
        private System.Windows.Forms.NumericUpDown numericUpDownStartNum;
        private System.Windows.Forms.Button buttonResume;
        private System.Windows.Forms.TextBox textBoxLog;
    }
}

