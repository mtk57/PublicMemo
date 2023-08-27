namespace Test00Form
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
            this.testUserButtonControl1 = new TestUserControl.TestUserButtonControl();
            this.test00CustomButton1 = new CustomTextBox.Test00CustomButton();
            this.test00ButtonEx1 = new Test00ControlAttributeTest.Test00ButtonEx();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // testUserButtonControl1
            // 
            this.testUserButtonControl1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.testUserButtonControl1.Location = new System.Drawing.Point(73, 297);
            this.testUserButtonControl1.Name = "testUserButtonControl1";
            this.testUserButtonControl1.Size = new System.Drawing.Size(292, 72);
            this.testUserButtonControl1.TabIndex = 3;
            this.testUserButtonControl1.Text = "test3(&F)";
            // 
            // test00CustomButton1
            // 
            this.test00CustomButton1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.test00CustomButton1.Location = new System.Drawing.Point(89, 171);
            this.test00CustomButton1.Name = "test00CustomButton1";
            this.test00CustomButton1.Size = new System.Drawing.Size(241, 58);
            this.test00CustomButton1.TabIndex = 1;
            this.test00CustomButton1.Text = "test1(&F)";
            this.test00CustomButton1.UseVisualStyleBackColor = true;
            // 
            // test00ButtonEx1
            // 
            this.test00ButtonEx1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.test00ButtonEx1.Location = new System.Drawing.Point(89, 235);
            this.test00ButtonEx1.Name = "test00ButtonEx1";
            this.test00ButtonEx1.Size = new System.Drawing.Size(241, 56);
            this.test00ButtonEx1.TabIndex = 2;
            this.test00ButtonEx1.Text = "test2(&F)";
            this.test00ButtonEx1.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.button1.Location = new System.Drawing.Point(89, 22);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(184, 66);
            this.button1.TabIndex = 0;
            this.button1.Text = "test0(&F)";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.test00ButtonEx1);
            this.Controls.Add(this.test00CustomButton1);
            this.Controls.Add(this.testUserButtonControl1);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private TestUserControl.TestUserButtonControl testUserButtonControl1;
        private CustomTextBox.Test00CustomButton test00CustomButton1;
        private Test00ControlAttributeTest.Test00ButtonEx test00ButtonEx1;
        private System.Windows.Forms.Button button1;
    }
}

