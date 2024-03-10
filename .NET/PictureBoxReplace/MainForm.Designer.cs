namespace PictureBoxReplace
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
        protected override void Dispose ( bool disposing )
        {
            if ( disposing && ( components != null ) )
            {
                components.Dispose();
            }
            base.Dispose( disposing );
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent ()
        {
            this.components = new System.ComponentModel.Container();
            this.gbBefore = new System.Windows.Forms.GroupBox();
            this.pbSelectImage = new System.Windows.Forms.PictureBox();
            this.tbInImageData = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbInPictureBoxName = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnInResxPathRef = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.tbInResxPath = new System.Windows.Forms.TextBox();
            this.gbAfter = new System.Windows.Forms.GroupBox();
            this.pbReplaceImage = new System.Windows.Forms.PictureBox();
            this.btnOutTargetDirPathRef = new System.Windows.Forms.Button();
            this.tbOutTargetDirPath = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cbOutCreateResxBak = new System.Windows.Forms.CheckBox();
            this.btnOutReplaceImageFilePathRef = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.tbOutReplaceImageFilePath = new System.Windows.Forms.TextBox();
            this.btnReplace = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.btnConfirmResult = new System.Windows.Forms.Button();
            this.lblPictureBoxCount = new System.Windows.Forms.Label();
            this.lblResxCount = new System.Windows.Forms.Label();
            this.gbBefore.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbSelectImage)).BeginInit();
            this.gbAfter.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbReplaceImage)).BeginInit();
            this.SuspendLayout();
            // 
            // gbBefore
            // 
            this.gbBefore.BackColor = System.Drawing.Color.LightSteelBlue;
            this.gbBefore.Controls.Add(this.lblPictureBoxCount);
            this.gbBefore.Controls.Add(this.pbSelectImage);
            this.gbBefore.Controls.Add(this.tbInImageData);
            this.gbBefore.Controls.Add(this.label3);
            this.gbBefore.Controls.Add(this.cmbInPictureBoxName);
            this.gbBefore.Controls.Add(this.label2);
            this.gbBefore.Controls.Add(this.btnInResxPathRef);
            this.gbBefore.Controls.Add(this.label1);
            this.gbBefore.Controls.Add(this.tbInResxPath);
            this.gbBefore.Location = new System.Drawing.Point(39, 40);
            this.gbBefore.Name = "gbBefore";
            this.gbBefore.Size = new System.Drawing.Size(1160, 473);
            this.gbBefore.TabIndex = 1;
            this.gbBefore.TabStop = false;
            this.gbBefore.Text = "■変更前情報";
            // 
            // pbSelectImage
            // 
            this.pbSelectImage.Location = new System.Drawing.Point(950, 137);
            this.pbSelectImage.Name = "pbSelectImage";
            this.pbSelectImage.Size = new System.Drawing.Size(150, 150);
            this.pbSelectImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pbSelectImage.TabIndex = 7;
            this.pbSelectImage.TabStop = false;
            // 
            // tbInImageData
            // 
            this.tbInImageData.Location = new System.Drawing.Point(54, 305);
            this.tbInImageData.Multiline = true;
            this.tbInImageData.Name = "tbInImageData";
            this.tbInImageData.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tbInImageData.Size = new System.Drawing.Size(1046, 137);
            this.tbInImageData.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(50, 273);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(224, 24);
            this.label3.TabIndex = 6;
            this.label3.Text = "③ 画像データ  [必須]";
            // 
            // cmbInPictureBoxName
            // 
            this.cmbInPictureBoxName.FormattingEnabled = true;
            this.cmbInPictureBoxName.Location = new System.Drawing.Point(45, 164);
            this.cmbInPictureBoxName.Name = "cmbInPictureBoxName";
            this.cmbInPictureBoxName.Size = new System.Drawing.Size(855, 32);
            this.cmbInPictureBoxName.TabIndex = 2;
            this.cmbInPictureBoxName.SelectedIndexChanged += new System.EventHandler(this.CmbInPictureBoxName_SelectedIndexChanged);
            this.cmbInPictureBoxName.Leave += new System.EventHandler(this.CmbInPictureBoxName_Leave);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(50, 137);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(219, 24);
            this.label2.TabIndex = 4;
            this.label2.Text = "② PictureBoxの名前";
            // 
            // btnInResxPathRef
            // 
            this.btnInResxPathRef.Location = new System.Drawing.Point(1029, 77);
            this.btnInResxPathRef.Name = "btnInResxPathRef";
            this.btnInResxPathRef.Size = new System.Drawing.Size(87, 46);
            this.btnInResxPathRef.TabIndex = 1;
            this.btnInResxPathRef.Text = "参照";
            this.btnInResxPathRef.UseVisualStyleBackColor = true;
            this.btnInResxPathRef.Click += new System.EventHandler(this.BtnInResxPathRef_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(50, 49);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(341, 24);
            this.label1.TabIndex = 1;
            this.label1.Text = "① resx ファイルパス (Drag Drop可)";
            // 
            // tbInResxPath
            // 
            this.tbInResxPath.AllowDrop = true;
            this.tbInResxPath.Location = new System.Drawing.Point(45, 82);
            this.tbInResxPath.Name = "tbInResxPath";
            this.tbInResxPath.Size = new System.Drawing.Size(961, 31);
            this.tbInResxPath.TabIndex = 0;
            this.tbInResxPath.DragDrop += new System.Windows.Forms.DragEventHandler(this.TbInResxPath_DragDrop);
            this.tbInResxPath.DragEnter += new System.Windows.Forms.DragEventHandler(this.TbInResxPath_DragEnter);
            this.tbInResxPath.Leave += new System.EventHandler(this.TbInResxPath_Leave);
            // 
            // gbAfter
            // 
            this.gbAfter.BackColor = System.Drawing.Color.Thistle;
            this.gbAfter.Controls.Add(this.lblResxCount);
            this.gbAfter.Controls.Add(this.pbReplaceImage);
            this.gbAfter.Controls.Add(this.btnOutTargetDirPathRef);
            this.gbAfter.Controls.Add(this.tbOutTargetDirPath);
            this.gbAfter.Controls.Add(this.label4);
            this.gbAfter.Controls.Add(this.cbOutCreateResxBak);
            this.gbAfter.Controls.Add(this.btnOutReplaceImageFilePathRef);
            this.gbAfter.Controls.Add(this.label6);
            this.gbAfter.Controls.Add(this.tbOutReplaceImageFilePath);
            this.gbAfter.Location = new System.Drawing.Point(39, 579);
            this.gbAfter.Name = "gbAfter";
            this.gbAfter.Size = new System.Drawing.Size(1160, 497);
            this.gbAfter.TabIndex = 2;
            this.gbAfter.TabStop = false;
            this.gbAfter.Text = "■変更後情報";
            // 
            // pbReplaceImage
            // 
            this.pbReplaceImage.Location = new System.Drawing.Point(950, 138);
            this.pbReplaceImage.Name = "pbReplaceImage";
            this.pbReplaceImage.Size = new System.Drawing.Size(150, 150);
            this.pbReplaceImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pbReplaceImage.TabIndex = 9;
            this.pbReplaceImage.TabStop = false;
            // 
            // btnOutTargetDirPathRef
            // 
            this.btnOutTargetDirPathRef.Location = new System.Drawing.Point(1029, 335);
            this.btnOutTargetDirPathRef.Name = "btnOutTargetDirPathRef";
            this.btnOutTargetDirPathRef.Size = new System.Drawing.Size(87, 46);
            this.btnOutTargetDirPathRef.TabIndex = 7;
            this.btnOutTargetDirPathRef.Text = "参照";
            this.btnOutTargetDirPathRef.UseVisualStyleBackColor = true;
            this.btnOutTargetDirPathRef.Click += new System.EventHandler(this.BtnOutTargetDirPathRef_Click);
            // 
            // tbOutTargetDirPath
            // 
            this.tbOutTargetDirPath.AllowDrop = true;
            this.tbOutTargetDirPath.Location = new System.Drawing.Point(45, 340);
            this.tbOutTargetDirPath.Name = "tbOutTargetDirPath";
            this.tbOutTargetDirPath.Size = new System.Drawing.Size(961, 31);
            this.tbOutTargetDirPath.TabIndex = 6;
            this.tbOutTargetDirPath.DragDrop += new System.Windows.Forms.DragEventHandler(this.TbOutTargetDirPath_DragDrop);
            this.tbOutTargetDirPath.DragEnter += new System.Windows.Forms.DragEventHandler(this.TbOutTargetDirPath_DragEnter);
            this.tbOutTargetDirPath.Leave += new System.EventHandler(this.TbOutTargetDirPath_Leave);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(50, 306);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(468, 24);
            this.label4.TabIndex = 4;
            this.label4.Text = "② 変換対象フォルダパス (Drag Drop可)  [必須]";
            // 
            // cbOutCreateResxBak
            // 
            this.cbOutCreateResxBak.AutoSize = true;
            this.cbOutCreateResxBak.Checked = true;
            this.cbOutCreateResxBak.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbOutCreateResxBak.Location = new System.Drawing.Point(54, 442);
            this.cbOutCreateResxBak.Name = "cbOutCreateResxBak";
            this.cbOutCreateResxBak.Size = new System.Drawing.Size(309, 28);
            this.cbOutCreateResxBak.TabIndex = 8;
            this.cbOutCreateResxBak.Text = "resxのバックアップを作成する";
            this.toolTip1.SetToolTip(this.cbOutCreateResxBak, "末尾に\".bak\"を付けます。\r\n同名ファイルが存在した場合は末尾に1からの連番を付けます。(例:hoge.resx.bak1)");
            this.cbOutCreateResxBak.UseVisualStyleBackColor = true;
            // 
            // btnOutReplaceImageFilePathRef
            // 
            this.btnOutReplaceImageFilePathRef.Location = new System.Drawing.Point(1029, 77);
            this.btnOutReplaceImageFilePathRef.Name = "btnOutReplaceImageFilePathRef";
            this.btnOutReplaceImageFilePathRef.Size = new System.Drawing.Size(87, 46);
            this.btnOutReplaceImageFilePathRef.TabIndex = 5;
            this.btnOutReplaceImageFilePathRef.Text = "参照";
            this.btnOutReplaceImageFilePathRef.UseVisualStyleBackColor = true;
            this.btnOutReplaceImageFilePathRef.Click += new System.EventHandler(this.BtnOutReplaceImageFilePathRef_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(50, 49);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(490, 24);
            this.label6.TabIndex = 1;
            this.label6.Text = "① 変換後画像ファイルパス (Drag Drop可)  [必須]";
            // 
            // tbOutReplaceImageFilePath
            // 
            this.tbOutReplaceImageFilePath.AllowDrop = true;
            this.tbOutReplaceImageFilePath.Location = new System.Drawing.Point(45, 82);
            this.tbOutReplaceImageFilePath.Name = "tbOutReplaceImageFilePath";
            this.tbOutReplaceImageFilePath.Size = new System.Drawing.Size(961, 31);
            this.tbOutReplaceImageFilePath.TabIndex = 4;
            this.tbOutReplaceImageFilePath.DragDrop += new System.Windows.Forms.DragEventHandler(this.TbOutReplaceImageFilePath_DragDrop);
            this.tbOutReplaceImageFilePath.DragEnter += new System.Windows.Forms.DragEventHandler(this.TbOutReplaceImageFilePath_DragEnter);
            this.tbOutReplaceImageFilePath.Leave += new System.EventHandler(this.TbOutReplaceImageFilePath_Leave);
            // 
            // btnReplace
            // 
            this.btnReplace.Location = new System.Drawing.Point(567, 1119);
            this.btnReplace.Name = "btnReplace";
            this.btnReplace.Size = new System.Drawing.Size(146, 57);
            this.btnReplace.TabIndex = 9;
            this.btnReplace.Text = "変換";
            this.btnReplace.UseVisualStyleBackColor = true;
            this.btnReplace.Click += new System.EventHandler(this.BtnReplace_Click);
            // 
            // btnConfirmResult
            // 
            this.btnConfirmResult.Location = new System.Drawing.Point(1000, 1119);
            this.btnConfirmResult.Name = "btnConfirmResult";
            this.btnConfirmResult.Size = new System.Drawing.Size(199, 57);
            this.btnConfirmResult.TabIndex = 10;
            this.btnConfirmResult.Text = "結果確認";
            this.btnConfirmResult.UseVisualStyleBackColor = true;
            this.btnConfirmResult.Click += new System.EventHandler(this.BtnConfirmResult_Click);
            // 
            // lblPictureBoxCount
            // 
            this.lblPictureBoxCount.AutoSize = true;
            this.lblPictureBoxCount.Location = new System.Drawing.Point(49, 205);
            this.lblPictureBoxCount.Name = "lblPictureBoxCount";
            this.lblPictureBoxCount.Size = new System.Drawing.Size(0, 24);
            this.lblPictureBoxCount.TabIndex = 8;
            // 
            // lblResxCount
            // 
            this.lblResxCount.AutoSize = true;
            this.lblResxCount.Location = new System.Drawing.Point(49, 384);
            this.lblResxCount.Name = "lblResxCount";
            this.lblResxCount.Size = new System.Drawing.Size(0, 24);
            this.lblResxCount.TabIndex = 10;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1230, 1223);
            this.Controls.Add(this.btnConfirmResult);
            this.Controls.Add(this.btnReplace);
            this.Controls.Add(this.gbAfter);
            this.Controls.Add(this.gbBefore);
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.Text = "PictureBoxReplace";
            this.gbBefore.ResumeLayout(false);
            this.gbBefore.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbSelectImage)).EndInit();
            this.gbAfter.ResumeLayout(false);
            this.gbAfter.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbReplaceImage)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox gbBefore;
        private System.Windows.Forms.TextBox tbInImageData;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmbInPictureBoxName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnInResxPathRef;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbInResxPath;
        private System.Windows.Forms.GroupBox gbAfter;
        private System.Windows.Forms.CheckBox cbOutCreateResxBak;
        private System.Windows.Forms.Button btnOutReplaceImageFilePathRef;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox tbOutReplaceImageFilePath;
        private System.Windows.Forms.Button btnReplace;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button btnOutTargetDirPathRef;
        private System.Windows.Forms.TextBox tbOutTargetDirPath;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.PictureBox pbSelectImage;
        private System.Windows.Forms.PictureBox pbReplaceImage;
        private System.Windows.Forms.Button btnConfirmResult;
        private System.Windows.Forms.Label lblPictureBoxCount;
        private System.Windows.Forms.Label lblResxCount;
    }
}

