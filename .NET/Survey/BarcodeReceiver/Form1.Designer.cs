namespace BarcodeReceiver
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
            this.textBoxBarcodeData = new System.Windows.Forms.TextBox();
            this.listBoxBarcodes = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // textBoxBarcodeData
            // 
            this.textBoxBarcodeData.Location = new System.Drawing.Point(66, 36);
            this.textBoxBarcodeData.Multiline = true;
            this.textBoxBarcodeData.Name = "textBoxBarcodeData";
            this.textBoxBarcodeData.Size = new System.Drawing.Size(359, 47);
            this.textBoxBarcodeData.TabIndex = 0;
            // 
            // listBoxBarcodes
            // 
            this.listBoxBarcodes.FormattingEnabled = true;
            this.listBoxBarcodes.ItemHeight = 12;
            this.listBoxBarcodes.Location = new System.Drawing.Point(66, 119);
            this.listBoxBarcodes.Name = "listBoxBarcodes";
            this.listBoxBarcodes.Size = new System.Drawing.Size(359, 136);
            this.listBoxBarcodes.TabIndex = 1;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(539, 372);
            this.Controls.Add(this.listBoxBarcodes);
            this.Controls.Add(this.textBoxBarcodeData);
            this.KeyPreview = true;
            this.Name = "Form1";
            this.Text = "Barcode Receiver";
            this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Form1_KeyPress);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxBarcodeData;
        private System.Windows.Forms.ListBox listBoxBarcodes;
    }
}

