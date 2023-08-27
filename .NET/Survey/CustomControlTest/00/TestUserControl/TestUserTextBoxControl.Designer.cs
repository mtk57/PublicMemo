namespace TestUserControl
{
    partial class TestUserTextBoxControl
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

        #region コンポーネント デザイナーで生成されたコード

        /// <summary> 
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を 
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.test00TextBoxEx1 = new Test00ControlAttributeTest.Test00TextBoxEx();
            this.SuspendLayout();
            // 
            // test00TextBoxEx1
            // 
            this.test00TextBoxEx1.Location = new System.Drawing.Point(61, 60);
            this.test00TextBoxEx1.Name = "test00TextBoxEx1";
            this.test00TextBoxEx1.Size = new System.Drawing.Size(100, 31);
            this.test00TextBoxEx1.TabIndex = 0;
            this.test00TextBoxEx1.TestBrowsable = false;
            this.test00TextBoxEx1.TestCategory = System.Drawing.Color.Empty;
            this.test00TextBoxEx1.TestDefaultValue2 = System.Drawing.Color.Empty;
            this.test00TextBoxEx1.TestDescription = null;
            this.test00TextBoxEx1.TestDesignerSerializationVisibility_Visible = false;
            // 
            // TestUserTextBoxControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.test00TextBoxEx1);
            this.Name = "TestUserTextBoxControl";
            this.Size = new System.Drawing.Size(312, 167);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Test00ControlAttributeTest.Test00TextBoxEx test00TextBoxEx1;
    }
}
