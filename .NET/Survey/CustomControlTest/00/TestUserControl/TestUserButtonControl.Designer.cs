namespace TestUserControl
{
    partial class TestUserButtonControl
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
            this.test00ButtonEx1 = new Test00ControlAttributeTest.Test00ButtonEx();
            this.SuspendLayout();
            // 
            // test00ButtonEx1
            // 
            this.test00ButtonEx1.Location = new System.Drawing.Point(14, 3);
            this.test00ButtonEx1.Name = "test00ButtonEx1";
            this.test00ButtonEx1.Size = new System.Drawing.Size(257, 53);
            this.test00ButtonEx1.TabIndex = 0;
            this.test00ButtonEx1.Text = "test00ButtonEx1";
            this.test00ButtonEx1.UseVisualStyleBackColor = true;
            this.test00ButtonEx1.UseMnemonic = true;
            // 
            // TestUserButtonControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.test00ButtonEx1);
            this.Name = "TestUserButtonControl";
            this.Size = new System.Drawing.Size(292, 72);
            this.ResumeLayout(false);

        }

        #endregion

        private Test00ControlAttributeTest.Test00ButtonEx test00ButtonEx1;
    }
}
