namespace TestHttps
{
    partial class FormMain
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if(disposing && (components != null))
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
            this.textBoxURL = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBoxTLS = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonGET = new System.Windows.Forms.Button();
            this.buttonPOST = new System.Windows.Forms.Button();
            this.textBoxParam1 = new System.Windows.Forms.TextBox();
            this.textBoxParam2 = new System.Windows.Forms.TextBox();
            this.textBoxValue2 = new System.Windows.Forms.TextBox();
            this.textBoxValue1 = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBoxURL
            // 
            this.textBoxURL.Location = new System.Drawing.Point(12, 41);
            this.textBoxURL.Name = "textBoxURL";
            this.textBoxURL.Size = new System.Drawing.Size(368, 19);
            this.textBoxURL.TabIndex = 0;
            this.textBoxURL.Text = "https://localhost:4433";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(27, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "URL";
            // 
            // comboBoxTLS
            // 
            this.comboBoxTLS.FormattingEnabled = true;
            this.comboBoxTLS.Items.AddRange(new object[] {
            "v1",
            "v1.1",
            "v1.2"});
            this.comboBoxTLS.Location = new System.Drawing.Point(14, 97);
            this.comboBoxTLS.Name = "comboBoxTLS";
            this.comboBoxTLS.Size = new System.Drawing.Size(74, 20);
            this.comboBoxTLS.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 82);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(45, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "TLS ver";
            // 
            // buttonGET
            // 
            this.buttonGET.Location = new System.Drawing.Point(14, 142);
            this.buttonGET.Name = "buttonGET";
            this.buttonGET.Size = new System.Drawing.Size(75, 23);
            this.buttonGET.TabIndex = 4;
            this.buttonGET.Text = "GET";
            this.buttonGET.UseVisualStyleBackColor = true;
            this.buttonGET.Click += new System.EventHandler(this.buttonGET_Click);
            // 
            // buttonPOST
            // 
            this.buttonPOST.Location = new System.Drawing.Point(224, 142);
            this.buttonPOST.Name = "buttonPOST";
            this.buttonPOST.Size = new System.Drawing.Size(75, 23);
            this.buttonPOST.TabIndex = 5;
            this.buttonPOST.Text = "POST";
            this.buttonPOST.UseVisualStyleBackColor = true;
            this.buttonPOST.Click += new System.EventHandler(this.buttonPOST_Click);
            // 
            // textBoxParam1
            // 
            this.textBoxParam1.Location = new System.Drawing.Point(6, 18);
            this.textBoxParam1.Name = "textBoxParam1";
            this.textBoxParam1.Size = new System.Drawing.Size(148, 19);
            this.textBoxParam1.TabIndex = 6;
            // 
            // textBoxParam2
            // 
            this.textBoxParam2.Location = new System.Drawing.Point(6, 43);
            this.textBoxParam2.Name = "textBoxParam2";
            this.textBoxParam2.Size = new System.Drawing.Size(148, 19);
            this.textBoxParam2.TabIndex = 7;
            // 
            // textBoxValue2
            // 
            this.textBoxValue2.Location = new System.Drawing.Point(6, 43);
            this.textBoxValue2.Name = "textBoxValue2";
            this.textBoxValue2.Size = new System.Drawing.Size(148, 19);
            this.textBoxValue2.TabIndex = 9;
            // 
            // textBoxValue1
            // 
            this.textBoxValue1.Location = new System.Drawing.Point(6, 18);
            this.textBoxValue1.Name = "textBoxValue1";
            this.textBoxValue1.Size = new System.Drawing.Size(148, 19);
            this.textBoxValue1.TabIndex = 8;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBoxParam1);
            this.groupBox1.Controls.Add(this.textBoxParam2);
            this.groupBox1.Location = new System.Drawing.Point(14, 184);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(174, 83);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Param";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.textBoxValue1);
            this.groupBox2.Controls.Add(this.textBoxValue2);
            this.groupBox2.Location = new System.Drawing.Point(224, 184);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(184, 83);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Value";
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(435, 289);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.buttonPOST);
            this.Controls.Add(this.buttonGET);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.comboBoxTLS);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxURL);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "FormMain";
            this.Text = "HTTPS Test";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxURL;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBoxTLS;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button buttonGET;
        private System.Windows.Forms.Button buttonPOST;
        private System.Windows.Forms.TextBox textBoxParam1;
        private System.Windows.Forms.TextBox textBoxParam2;
        private System.Windows.Forms.TextBox textBoxValue2;
        private System.Windows.Forms.TextBox textBoxValue1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
    }
}

