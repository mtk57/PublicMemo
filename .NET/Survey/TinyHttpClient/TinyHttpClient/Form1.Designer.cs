namespace TinyHttpClient
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
            this.components = new System.ComponentModel.Container();
            this.textBoxURL = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBoxTLS = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonGET = new System.Windows.Forms.Button();
            this.buttonPOST = new System.Windows.Forms.Button();
            this.textBoxValueName = new System.Windows.Forms.TextBox();
            this.textBoxValueId = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.buttonPUT = new System.Windows.Forms.Button();
            this.buttonDELETE = new System.Windows.Forms.Button();
            this.radioButtonHttp = new System.Windows.Forms.RadioButton();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.radioButtonHttps = new System.Windows.Forms.RadioButton();
            this.buttonDefaultURL = new System.Windows.Forms.Button();
            this.buttonDefaultParam = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.textBox_Log = new System.Windows.Forms.TextBox();
            this.toolTip2 = new System.Windows.Forms.ToolTip(this.components);
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBoxURL
            // 
            this.textBoxURL.Location = new System.Drawing.Point(41, 63);
            this.textBoxURL.Margin = new System.Windows.Forms.Padding(6);
            this.textBoxURL.Name = "textBoxURL";
            this.textBoxURL.Size = new System.Drawing.Size(793, 31);
            this.textBoxURL.TabIndex = 0;
            this.toolTip1.SetToolTip(this.textBoxURL, "GET, DELETEの場合は末尾にIDを指定すること");
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(37, 33);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 24);
            this.label1.TabIndex = 1;
            this.label1.Text = "URL";
            // 
            // comboBoxTLS
            // 
            this.comboBoxTLS.FormattingEnabled = true;
            this.comboBoxTLS.Location = new System.Drawing.Point(363, 169);
            this.comboBoxTLS.Margin = new System.Windows.Forms.Padding(6);
            this.comboBoxTLS.Name = "comboBoxTLS";
            this.comboBoxTLS.Size = new System.Drawing.Size(156, 32);
            this.comboBoxTLS.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(359, 139);
            this.label2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 24);
            this.label2.TabIndex = 3;
            this.label2.Text = "TLS ver";
            // 
            // buttonGET
            // 
            this.buttonGET.Location = new System.Drawing.Point(233, 266);
            this.buttonGET.Margin = new System.Windows.Forms.Padding(6);
            this.buttonGET.Name = "buttonGET";
            this.buttonGET.Size = new System.Drawing.Size(162, 46);
            this.buttonGET.TabIndex = 4;
            this.buttonGET.Text = "GET (R)";
            this.buttonGET.UseVisualStyleBackColor = true;
            this.buttonGET.Click += new System.EventHandler(this.buttonGET_Click);
            // 
            // buttonPOST
            // 
            this.buttonPOST.Location = new System.Drawing.Point(41, 266);
            this.buttonPOST.Margin = new System.Windows.Forms.Padding(6);
            this.buttonPOST.Name = "buttonPOST";
            this.buttonPOST.Size = new System.Drawing.Size(162, 46);
            this.buttonPOST.TabIndex = 5;
            this.buttonPOST.Text = "POST (C)";
            this.buttonPOST.UseVisualStyleBackColor = true;
            this.buttonPOST.Click += new System.EventHandler(this.buttonPOST_Click);
            // 
            // textBoxValueName
            // 
            this.textBoxValueName.Location = new System.Drawing.Point(131, 107);
            this.textBoxValueName.Margin = new System.Windows.Forms.Padding(6);
            this.textBoxValueName.Name = "textBoxValueName";
            this.textBoxValueName.Size = new System.Drawing.Size(256, 31);
            this.textBoxValueName.TabIndex = 9;
            this.toolTip1.SetToolTip(this.textBoxValueName, "何でも可");
            // 
            // textBoxValueId
            // 
            this.textBoxValueId.Location = new System.Drawing.Point(131, 57);
            this.textBoxValueId.Margin = new System.Windows.Forms.Padding(6);
            this.textBoxValueId.Name = "textBoxValueId";
            this.textBoxValueId.Size = new System.Drawing.Size(256, 31);
            this.textBoxValueId.TabIndex = 8;
            this.toolTip1.SetToolTip(this.textBoxValueId, "1以上の整数");
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.textBoxValueId);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.textBoxValueName);
            this.groupBox1.Location = new System.Drawing.Point(28, 343);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox1.Size = new System.Drawing.Size(418, 166);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Param";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(29, 107);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(63, 24);
            this.label6.TabIndex = 16;
            this.label6.Text = "name";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(29, 60);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(27, 24);
            this.label5.TabIndex = 16;
            this.label5.Text = "id";
            // 
            // buttonPUT
            // 
            this.buttonPUT.Location = new System.Drawing.Point(443, 266);
            this.buttonPUT.Margin = new System.Windows.Forms.Padding(6);
            this.buttonPUT.Name = "buttonPUT";
            this.buttonPUT.Size = new System.Drawing.Size(162, 46);
            this.buttonPUT.TabIndex = 12;
            this.buttonPUT.Text = "PUT (U)";
            this.buttonPUT.UseVisualStyleBackColor = true;
            this.buttonPUT.Click += new System.EventHandler(this.buttonPUT_Click);
            // 
            // buttonDELETE
            // 
            this.buttonDELETE.Location = new System.Drawing.Point(650, 266);
            this.buttonDELETE.Margin = new System.Windows.Forms.Padding(6);
            this.buttonDELETE.Name = "buttonDELETE";
            this.buttonDELETE.Size = new System.Drawing.Size(162, 46);
            this.buttonDELETE.TabIndex = 13;
            this.buttonDELETE.Text = "DELETE (D)";
            this.buttonDELETE.UseVisualStyleBackColor = true;
            this.buttonDELETE.Click += new System.EventHandler(this.buttonDELETE_Click);
            // 
            // radioButtonHttp
            // 
            this.radioButtonHttp.AutoSize = true;
            this.radioButtonHttp.Checked = true;
            this.radioButtonHttp.Location = new System.Drawing.Point(20, 44);
            this.radioButtonHttp.Name = "radioButtonHttp";
            this.radioButtonHttp.Size = new System.Drawing.Size(81, 28);
            this.radioButtonHttp.TabIndex = 14;
            this.radioButtonHttp.TabStop = true;
            this.radioButtonHttp.Text = "http";
            this.radioButtonHttp.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.radioButtonHttps);
            this.groupBox2.Controls.Add(this.radioButtonHttp);
            this.groupBox2.Location = new System.Drawing.Point(41, 115);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(271, 100);
            this.groupBox2.TabIndex = 15;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Scheme";
            // 
            // radioButtonHttps
            // 
            this.radioButtonHttps.AutoSize = true;
            this.radioButtonHttps.Location = new System.Drawing.Point(130, 44);
            this.radioButtonHttps.Name = "radioButtonHttps";
            this.radioButtonHttps.Size = new System.Drawing.Size(92, 28);
            this.radioButtonHttps.TabIndex = 15;
            this.radioButtonHttps.Text = "https";
            this.radioButtonHttps.UseVisualStyleBackColor = true;
            // 
            // buttonDefaultURL
            // 
            this.buttonDefaultURL.Location = new System.Drawing.Point(583, 155);
            this.buttonDefaultURL.Margin = new System.Windows.Forms.Padding(6);
            this.buttonDefaultURL.Name = "buttonDefaultURL";
            this.buttonDefaultURL.Size = new System.Drawing.Size(203, 46);
            this.buttonDefaultURL.TabIndex = 16;
            this.buttonDefaultURL.Text = "Default URL";
            this.buttonDefaultURL.UseVisualStyleBackColor = true;
            this.buttonDefaultURL.Click += new System.EventHandler(this.buttonDefaultURL_Click);
            // 
            // buttonDefaultParam
            // 
            this.buttonDefaultParam.Location = new System.Drawing.Point(583, 381);
            this.buttonDefaultParam.Margin = new System.Windows.Forms.Padding(6);
            this.buttonDefaultParam.Name = "buttonDefaultParam";
            this.buttonDefaultParam.Size = new System.Drawing.Size(203, 46);
            this.buttonDefaultParam.TabIndex = 17;
            this.buttonDefaultParam.Text = "Default Param";
            this.buttonDefaultParam.UseVisualStyleBackColor = true;
            this.buttonDefaultParam.Click += new System.EventHandler(this.buttonDefaultParam_Click);
            // 
            // textBox_Log
            // 
            this.textBox_Log.Location = new System.Drawing.Point(28, 518);
            this.textBox_Log.Multiline = true;
            this.textBox_Log.Name = "textBox_Log";
            this.textBox_Log.ReadOnly = true;
            this.textBox_Log.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox_Log.Size = new System.Drawing.Size(784, 292);
            this.textBox_Log.TabIndex = 18;
            this.toolTip1.SetToolTip(this.textBox_Log, "Ex:123,456");
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(862, 834);
            this.Controls.Add(this.textBox_Log);
            this.Controls.Add(this.buttonDefaultParam);
            this.Controls.Add(this.buttonDefaultURL);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.buttonDELETE);
            this.Controls.Add(this.buttonPUT);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.buttonPOST);
            this.Controls.Add(this.buttonGET);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.comboBoxTLS);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxURL);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(6);
            this.MaximizeBox = false;
            this.Name = "FormMain";
            this.Text = "Tiny Http Client";
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
        private System.Windows.Forms.TextBox textBoxValueName;
        private System.Windows.Forms.TextBox textBoxValueId;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button buttonPUT;
        private System.Windows.Forms.Button buttonDELETE;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.RadioButton radioButtonHttp;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton radioButtonHttps;
        private System.Windows.Forms.Button buttonDefaultURL;
        private System.Windows.Forms.Button buttonDefaultParam;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.TextBox textBox_Log;
        private System.Windows.Forms.ToolTip toolTip2;
    }
}

