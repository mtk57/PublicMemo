namespace ComLibTestKicker
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
            this.buttonGetUserInfos = new System.Windows.Forms.Button();
            this.textBoxLog = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // buttonGetUserInfos
            // 
            this.buttonGetUserInfos.Location = new System.Drawing.Point(32, 32);
            this.buttonGetUserInfos.Name = "buttonGetUserInfos";
            this.buttonGetUserInfos.Size = new System.Drawing.Size(223, 50);
            this.buttonGetUserInfos.TabIndex = 0;
            this.buttonGetUserInfos.Text = "Get UserInfos";
            this.buttonGetUserInfos.UseVisualStyleBackColor = true;
            this.buttonGetUserInfos.Click += new System.EventHandler(this.buttonGetUserInfos_Click);
            // 
            // textBoxLog
            // 
            this.textBoxLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxLog.Location = new System.Drawing.Point(32, 101);
            this.textBoxLog.Multiline = true;
            this.textBoxLog.Name = "textBoxLog";
            this.textBoxLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxLog.Size = new System.Drawing.Size(804, 331);
            this.textBoxLog.TabIndex = 1;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(859, 459);
            this.Controls.Add(this.textBoxLog);
            this.Controls.Add(this.buttonGetUserInfos);
            this.Name = "Form1";
            this.Text = "ComLibTest Kicker";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonGetUserInfos;
        private System.Windows.Forms.TextBox textBoxLog;
    }
}

