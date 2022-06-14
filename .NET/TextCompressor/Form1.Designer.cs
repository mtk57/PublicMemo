namespace TextCompressor
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
            this.textBoxCompInDirPath = new System.Windows.Forms.TextBox();
            this.buttonCompInDirRef = new System.Windows.Forms.Button();
            this.textBoxCompExt = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.textBoxKeyword = new System.Windows.Forms.TextBox();
            this.buttonCompOutDirRef = new System.Windows.Forms.Button();
            this.textBoxCompOutDirPath = new System.Windows.Forms.TextBox();
            this.buttonRunComp = new System.Windows.Forms.Button();
            this.buttonDecompInFileRef = new System.Windows.Forms.Button();
            this.textBoxDecompInFilePath = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.buttonDecompOutDirRef = new System.Windows.Forms.Button();
            this.textBoxDecompOutDirPath = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.buttonRunDecomp = new System.Windows.Forms.Button();
            this.buttonDefault = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 137);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(110, 24);
            this.label1.TabIndex = 0;
            this.label1.Text = "Compress";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(29, 525);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(134, 24);
            this.label2.TabIndex = 1;
            this.label2.Text = "Decompress";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(51, 188);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(149, 24);
            this.label3.TabIndex = 2;
            this.label3.Text = "Input Dir Path";
            // 
            // textBoxCompInDirPath
            // 
            this.textBoxCompInDirPath.Location = new System.Drawing.Point(55, 215);
            this.textBoxCompInDirPath.Name = "textBoxCompInDirPath";
            this.textBoxCompInDirPath.Size = new System.Drawing.Size(723, 31);
            this.textBoxCompInDirPath.TabIndex = 3;
            // 
            // buttonCompInDirRef
            // 
            this.buttonCompInDirRef.Location = new System.Drawing.Point(797, 215);
            this.buttonCompInDirRef.Name = "buttonCompInDirRef";
            this.buttonCompInDirRef.Size = new System.Drawing.Size(56, 31);
            this.buttonCompInDirRef.TabIndex = 4;
            this.buttonCompInDirRef.Text = "Ref";
            this.buttonCompInDirRef.UseVisualStyleBackColor = true;
            this.buttonCompInDirRef.Click += new System.EventHandler(this.buttonCompInDirRef_Click);
            // 
            // textBoxCompExt
            // 
            this.textBoxCompExt.Location = new System.Drawing.Point(55, 287);
            this.textBoxCompExt.Name = "textBoxCompExt";
            this.textBoxCompExt.Size = new System.Drawing.Size(439, 31);
            this.textBoxCompExt.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(51, 260);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(288, 24);
            this.label4.TabIndex = 6;
            this.label4.Text = "Target Extension (Ex. cs|vb)";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(29, 38);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(95, 24);
            this.label5.TabIndex = 7;
            this.label5.Text = "Keyword";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(51, 354);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(322, 24);
            this.label6.TabIndex = 8;
            this.label6.Text = "Output Compress-File Dir Path";
            // 
            // textBoxKeyword
            // 
            this.textBoxKeyword.Location = new System.Drawing.Point(147, 38);
            this.textBoxKeyword.Name = "textBoxKeyword";
            this.textBoxKeyword.Size = new System.Drawing.Size(631, 31);
            this.textBoxKeyword.TabIndex = 9;
            // 
            // buttonCompOutDirRef
            // 
            this.buttonCompOutDirRef.Location = new System.Drawing.Point(797, 381);
            this.buttonCompOutDirRef.Name = "buttonCompOutDirRef";
            this.buttonCompOutDirRef.Size = new System.Drawing.Size(56, 31);
            this.buttonCompOutDirRef.TabIndex = 11;
            this.buttonCompOutDirRef.Text = "Ref";
            this.buttonCompOutDirRef.UseVisualStyleBackColor = true;
            this.buttonCompOutDirRef.Click += new System.EventHandler(this.buttonCompOutDirRef_Click);
            // 
            // textBoxCompOutDirPath
            // 
            this.textBoxCompOutDirPath.Location = new System.Drawing.Point(55, 381);
            this.textBoxCompOutDirPath.Name = "textBoxCompOutDirPath";
            this.textBoxCompOutDirPath.Size = new System.Drawing.Size(723, 31);
            this.textBoxCompOutDirPath.TabIndex = 10;
            // 
            // buttonRunComp
            // 
            this.buttonRunComp.Location = new System.Drawing.Point(295, 435);
            this.buttonRunComp.Name = "buttonRunComp";
            this.buttonRunComp.Size = new System.Drawing.Size(252, 36);
            this.buttonRunComp.TabIndex = 12;
            this.buttonRunComp.Text = "Run Compress";
            this.buttonRunComp.UseVisualStyleBackColor = true;
            this.buttonRunComp.Click += new System.EventHandler(this.buttonRunComp_Click);
            // 
            // buttonDecompInFileRef
            // 
            this.buttonDecompInFileRef.Location = new System.Drawing.Point(797, 592);
            this.buttonDecompInFileRef.Name = "buttonDecompInFileRef";
            this.buttonDecompInFileRef.Size = new System.Drawing.Size(56, 31);
            this.buttonDecompInFileRef.TabIndex = 15;
            this.buttonDecompInFileRef.Text = "Ref";
            this.buttonDecompInFileRef.UseVisualStyleBackColor = true;
            this.buttonDecompInFileRef.Click += new System.EventHandler(this.buttonDecompInFileRef_Click);
            // 
            // textBoxDecompInFilePath
            // 
            this.textBoxDecompInFilePath.Location = new System.Drawing.Point(55, 592);
            this.textBoxDecompInFilePath.Name = "textBoxDecompInFilePath";
            this.textBoxDecompInFilePath.Size = new System.Drawing.Size(723, 31);
            this.textBoxDecompInFilePath.TabIndex = 14;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(51, 565);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(267, 24);
            this.label7.TabIndex = 13;
            this.label7.Text = "Input Compress-File Path";
            // 
            // buttonDecompOutDirRef
            // 
            this.buttonDecompOutDirRef.Location = new System.Drawing.Point(797, 682);
            this.buttonDecompOutDirRef.Name = "buttonDecompOutDirRef";
            this.buttonDecompOutDirRef.Size = new System.Drawing.Size(56, 31);
            this.buttonDecompOutDirRef.TabIndex = 18;
            this.buttonDecompOutDirRef.Text = "Ref";
            this.buttonDecompOutDirRef.UseVisualStyleBackColor = true;
            this.buttonDecompOutDirRef.Click += new System.EventHandler(this.buttonDecompOutDirRef_Click);
            // 
            // textBoxDecompOutDirPath
            // 
            this.textBoxDecompOutDirPath.Location = new System.Drawing.Point(55, 682);
            this.textBoxDecompOutDirPath.Name = "textBoxDecompOutDirPath";
            this.textBoxDecompOutDirPath.Size = new System.Drawing.Size(723, 31);
            this.textBoxDecompOutDirPath.TabIndex = 17;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(51, 655);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(168, 24);
            this.label8.TabIndex = 16;
            this.label8.Text = "Output Dir Path";
            // 
            // buttonRunDecomp
            // 
            this.buttonRunDecomp.Location = new System.Drawing.Point(295, 750);
            this.buttonRunDecomp.Name = "buttonRunDecomp";
            this.buttonRunDecomp.Size = new System.Drawing.Size(252, 36);
            this.buttonRunDecomp.TabIndex = 19;
            this.buttonRunDecomp.Text = "Run Decompress";
            this.buttonRunDecomp.UseVisualStyleBackColor = true;
            this.buttonRunDecomp.Click += new System.EventHandler(this.buttonRunDecomp_Click);
            // 
            // buttonDefault
            // 
            this.buttonDefault.Location = new System.Drawing.Point(601, 114);
            this.buttonDefault.Name = "buttonDefault";
            this.buttonDefault.Size = new System.Drawing.Size(155, 32);
            this.buttonDefault.TabIndex = 20;
            this.buttonDefault.Text = "Default";
            this.buttonDefault.UseVisualStyleBackColor = true;
            this.buttonDefault.Click += new System.EventHandler(this.buttonDefault_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(875, 823);
            this.Controls.Add(this.buttonDefault);
            this.Controls.Add(this.buttonRunDecomp);
            this.Controls.Add(this.buttonDecompOutDirRef);
            this.Controls.Add(this.textBoxDecompOutDirPath);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.buttonDecompInFileRef);
            this.Controls.Add(this.textBoxDecompInFilePath);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.buttonRunComp);
            this.Controls.Add(this.buttonCompOutDirRef);
            this.Controls.Add(this.textBoxCompOutDirPath);
            this.Controls.Add(this.textBoxKeyword);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBoxCompExt);
            this.Controls.Add(this.buttonCompInDirRef);
            this.Controls.Add(this.textBoxCompInDirPath);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.MaximumSize = new System.Drawing.Size(901, 894);
            this.MinimumSize = new System.Drawing.Size(901, 894);
            this.Name = "Form1";
            this.Text = "Text Compressor";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxCompInDirPath;
        private System.Windows.Forms.Button buttonCompInDirRef;
        private System.Windows.Forms.TextBox textBoxCompExt;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBoxKeyword;
        private System.Windows.Forms.Button buttonCompOutDirRef;
        private System.Windows.Forms.TextBox textBoxCompOutDirPath;
        private System.Windows.Forms.Button buttonRunComp;
        private System.Windows.Forms.Button buttonDecompInFileRef;
        private System.Windows.Forms.TextBox textBoxDecompInFilePath;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button buttonDecompOutDirRef;
        private System.Windows.Forms.TextBox textBoxDecompOutDirPath;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button buttonRunDecomp;
        private System.Windows.Forms.Button buttonDefault;
    }
}

