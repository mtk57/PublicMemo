namespace BarcodeSender
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
            this.textBoxSendData = new System.Windows.Forms.TextBox();
            this.buttonSend = new System.Windows.Forms.Button();
            this.textBoxReceiveWindowCaption = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // textBoxSendData
            // 
            this.textBoxSendData.Location = new System.Drawing.Point(28, 89);
            this.textBoxSendData.Multiline = true;
            this.textBoxSendData.Name = "textBoxSendData";
            this.textBoxSendData.Size = new System.Drawing.Size(438, 65);
            this.textBoxSendData.TabIndex = 0;
            // 
            // buttonSend
            // 
            this.buttonSend.Location = new System.Drawing.Point(202, 177);
            this.buttonSend.Name = "buttonSend";
            this.buttonSend.Size = new System.Drawing.Size(88, 33);
            this.buttonSend.TabIndex = 1;
            this.buttonSend.Text = "Send";
            this.buttonSend.UseVisualStyleBackColor = true;
            this.buttonSend.Click += new System.EventHandler(this.buttonSend_Click);
            // 
            // textBoxReceiveWindowCaption
            // 
            this.textBoxReceiveWindowCaption.Location = new System.Drawing.Point(28, 34);
            this.textBoxReceiveWindowCaption.Name = "textBoxReceiveWindowCaption";
            this.textBoxReceiveWindowCaption.Size = new System.Drawing.Size(241, 19);
            this.textBoxReceiveWindowCaption.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(131, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "Receive Window Caption";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(28, 74);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(102, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "Send Barcode data";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(501, 241);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxReceiveWindowCaption);
            this.Controls.Add(this.buttonSend);
            this.Controls.Add(this.textBoxSendData);
            this.Name = "Form1";
            this.Text = "Barcode Sender";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxSendData;
        private System.Windows.Forms.Button buttonSend;
        private System.Windows.Forms.TextBox textBoxReceiveWindowCaption;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}

