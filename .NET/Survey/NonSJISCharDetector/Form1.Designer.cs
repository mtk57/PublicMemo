namespace NonSJISCharDetector
{
    partial class Form1
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
            this.textBoxDirPath = new System.Windows.Forms.TextBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.checkBoxReplaceSpace = new System.Windows.Forms.CheckBox();
            this.buttonRun = new System.Windows.Forms.Button();
            this.buttonRefDir = new System.Windows.Forms.Button();
            this.textBoxExt = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.checkBoxCreateBackup = new System.Windows.Forms.CheckBox();
            this.textBoxOutDirPath = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.buttonRefOutDir = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBoxDirPath
            // 
            this.textBoxDirPath.AllowDrop = true;
            this.textBoxDirPath.Location = new System.Drawing.Point(33, 43);
            this.textBoxDirPath.Name = "textBoxDirPath";
            this.textBoxDirPath.Size = new System.Drawing.Size(460, 19);
            this.textBoxDirPath.TabIndex = 0;
            this.textBoxDirPath.DragDrop += new System.Windows.Forms.DragEventHandler(this.textBoxDirPath_DragDrop);
            this.textBoxDirPath.DragEnter += new System.Windows.Forms.DragEventHandler(this.textBoxDirPath_DragEnter);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(33, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "Target Dir path";
            // 
            // checkBoxReplaceSpace
            // 
            this.checkBoxReplaceSpace.AutoSize = true;
            this.checkBoxReplaceSpace.Location = new System.Drawing.Point(35, 198);
            this.checkBoxReplaceSpace.Name = "checkBoxReplaceSpace";
            this.checkBoxReplaceSpace.Size = new System.Drawing.Size(106, 16);
            this.checkBoxReplaceSpace.TabIndex = 3;
            this.checkBoxReplaceSpace.Text = "Replace SPACE";
            this.checkBoxReplaceSpace.UseVisualStyleBackColor = true;
            // 
            // buttonRun
            // 
            this.buttonRun.Location = new System.Drawing.Point(214, 252);
            this.buttonRun.Name = "buttonRun";
            this.buttonRun.Size = new System.Drawing.Size(124, 23);
            this.buttonRun.TabIndex = 4;
            this.buttonRun.Text = "RUN";
            this.buttonRun.UseVisualStyleBackColor = true;
            this.buttonRun.Click += new System.EventHandler(this.buttonRun_Click);
            // 
            // buttonRefDir
            // 
            this.buttonRefDir.Location = new System.Drawing.Point(500, 38);
            this.buttonRefDir.Name = "buttonRefDir";
            this.buttonRefDir.Size = new System.Drawing.Size(52, 23);
            this.buttonRefDir.TabIndex = 5;
            this.buttonRefDir.Text = "ref";
            this.buttonRefDir.UseVisualStyleBackColor = true;
            this.buttonRefDir.Click += new System.EventHandler(this.buttonRefDir_Click);
            // 
            // textBoxExt
            // 
            this.textBoxExt.Location = new System.Drawing.Point(33, 94);
            this.textBoxExt.Name = "textBoxExt";
            this.textBoxExt.Size = new System.Drawing.Size(145, 19);
            this.textBoxExt.TabIndex = 6;
            this.textBoxExt.Text = "txt,frm,bas,cls,ctl";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(33, 79);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 12);
            this.label2.TabIndex = 7;
            this.label2.Text = "Target Extension";
            // 
            // checkBoxCreateBackup
            // 
            this.checkBoxCreateBackup.AutoSize = true;
            this.checkBoxCreateBackup.Location = new System.Drawing.Point(184, 198);
            this.checkBoxCreateBackup.Name = "checkBoxCreateBackup";
            this.checkBoxCreateBackup.Size = new System.Drawing.Size(100, 16);
            this.checkBoxCreateBackup.TabIndex = 8;
            this.checkBoxCreateBackup.Text = "Create Backup";
            this.checkBoxCreateBackup.UseVisualStyleBackColor = true;
            // 
            // textBoxOutDirPath
            // 
            this.textBoxOutDirPath.AllowDrop = true;
            this.textBoxOutDirPath.Location = new System.Drawing.Point(33, 148);
            this.textBoxOutDirPath.Name = "textBoxOutDirPath";
            this.textBoxOutDirPath.Size = new System.Drawing.Size(460, 19);
            this.textBoxOutDirPath.TabIndex = 9;
            this.textBoxOutDirPath.DragDrop += new System.Windows.Forms.DragEventHandler(this.textBoxOutDirPath_DragDrop);
            this.textBoxOutDirPath.DragEnter += new System.Windows.Forms.DragEventHandler(this.textBoxOutDirPath_DragEnter);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(33, 133);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(84, 12);
            this.label3.TabIndex = 10;
            this.label3.Text = "Output Dir path";
            // 
            // buttonRefOutDir
            // 
            this.buttonRefOutDir.Location = new System.Drawing.Point(500, 148);
            this.buttonRefOutDir.Name = "buttonRefOutDir";
            this.buttonRefOutDir.Size = new System.Drawing.Size(52, 23);
            this.buttonRefOutDir.TabIndex = 11;
            this.buttonRefOutDir.Text = "ref";
            this.buttonRefOutDir.UseVisualStyleBackColor = true;
            this.buttonRefOutDir.Click += new System.EventHandler(this.buttonRefOutDir_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(562, 299);
            this.Controls.Add(this.buttonRefOutDir);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBoxOutDirPath);
            this.Controls.Add(this.checkBoxCreateBackup);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBoxExt);
            this.Controls.Add(this.buttonRefDir);
            this.Controls.Add(this.buttonRun);
            this.Controls.Add(this.checkBoxReplaceSpace);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxDirPath);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxDirPath;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox checkBoxReplaceSpace;
        private System.Windows.Forms.Button buttonRun;
        private System.Windows.Forms.Button buttonRefDir;
        private System.Windows.Forms.TextBox textBoxExt;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox checkBoxCreateBackup;
        private System.Windows.Forms.TextBox textBoxOutDirPath;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button buttonRefOutDir;
    }
}

