namespace TabOrderHelper
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
            this.userControl11 = new MyControl.UserControl1();
            this.userControl12 = new MyControl.UserControl1();
            this.userControl13 = new MyControl.UserControl1();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.userControl14 = new MyControl.UserControl1();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // userControl11
            // 
            this.userControl11.Location = new System.Drawing.Point(67, 37);
            this.userControl11.Name = "userControl11";
            this.userControl11.Size = new System.Drawing.Size(215, 40);
            this.userControl11.TabIndex = 6;
            // 
            // userControl12
            // 
            this.userControl12.Enabled = false;
            this.userControl12.Location = new System.Drawing.Point(316, 37);
            this.userControl12.Name = "userControl12";
            this.userControl12.Size = new System.Drawing.Size(215, 40);
            this.userControl12.TabIndex = 7;
            // 
            // userControl13
            // 
            this.userControl13.Location = new System.Drawing.Point(572, 37);
            this.userControl13.Name = "userControl13";
            this.userControl13.Size = new System.Drawing.Size(215, 40);
            this.userControl13.TabIndex = 8;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButton2);
            this.groupBox1.Controls.Add(this.radioButton1);
            this.groupBox1.Controls.Add(this.userControl14);
            this.groupBox1.Location = new System.Drawing.Point(86, 173);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(513, 325);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "groupBox1";
            // 
            // userControl14
            // 
            this.userControl14.Location = new System.Drawing.Point(53, 77);
            this.userControl14.Name = "userControl14";
            this.userControl14.Size = new System.Drawing.Size(215, 40);
            this.userControl14.TabIndex = 10;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new System.Drawing.Point(62, 169);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(168, 28);
            this.radioButton1.TabIndex = 11;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "radioButton1";
            this.radioButton1.UseVisualStyleBackColor = true;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(62, 234);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(168, 28);
            this.radioButton2.TabIndex = 12;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "radioButton2";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1315, 857);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.userControl13);
            this.Controls.Add(this.userControl12);
            this.Controls.Add(this.userControl11);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private MyControl.UserControl1 userControl11;
        private MyControl.UserControl1 userControl12;
        private MyControl.UserControl1 userControl13;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.RadioButton radioButton1;
        private MyControl.UserControl1 userControl14;
    }
}

