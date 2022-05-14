namespace TinyHttpClient2
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
            this.buttonPostToken = new System.Windows.Forms.Button();
            this.textBoxToken = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonReload = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonPostToken
            // 
            this.buttonPostToken.Location = new System.Drawing.Point(59, 48);
            this.buttonPostToken.Name = "buttonPostToken";
            this.buttonPostToken.Size = new System.Drawing.Size(393, 39);
            this.buttonPostToken.TabIndex = 0;
            this.buttonPostToken.Text = "Post Token";
            this.buttonPostToken.UseVisualStyleBackColor = true;
            this.buttonPostToken.Click += new System.EventHandler(this.buttonPostToken_Click);
            // 
            // textBoxToken
            // 
            this.textBoxToken.Location = new System.Drawing.Point(132, 107);
            this.textBoxToken.Name = "textBoxToken";
            this.textBoxToken.Size = new System.Drawing.Size(320, 31);
            this.textBoxToken.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(55, 107);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 24);
            this.label1.TabIndex = 2;
            this.label1.Text = "Token";
            // 
            // buttonReload
            // 
            this.buttonReload.Location = new System.Drawing.Point(59, 332);
            this.buttonReload.Name = "buttonReload";
            this.buttonReload.Size = new System.Drawing.Size(210, 39);
            this.buttonReload.TabIndex = 3;
            this.buttonReload.Text = "Reload settings";
            this.buttonReload.UseVisualStyleBackColor = true;
            this.buttonReload.Click += new System.EventHandler(this.buttonReload_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(528, 399);
            this.Controls.Add(this.buttonReload);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxToken);
            this.Controls.Add(this.buttonPostToken);
            this.Name = "Form1";
            this.Text = "LibCOM_Kicker";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonPostToken;
        private System.Windows.Forms.TextBox textBoxToken;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonReload;
    }
}

