namespace TinyRestServer
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
            this.components = new System.ComponentModel.Container();
            this.button_Start = new System.Windows.Forms.Button();
            this.button_Stop = new System.Windows.Forms.Button();
            this.textBox_Port = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox_Url = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.textBox_Log = new System.Windows.Forms.TextBox();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.buttonDefaultURL = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.radioButtonHttps = new System.Windows.Forms.RadioButton();
            this.radioButtonHttp = new System.Windows.Forms.RadioButton();
            this.buttonAddCert = new System.Windows.Forms.Button();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // button_Start
            // 
            this.button_Start.Location = new System.Drawing.Point(12, 223);
            this.button_Start.Name = "button_Start";
            this.button_Start.Size = new System.Drawing.Size(204, 54);
            this.button_Start.TabIndex = 0;
            this.button_Start.Text = "Start";
            this.button_Start.UseVisualStyleBackColor = true;
            this.button_Start.Click += new System.EventHandler(this.button_Start_Click);
            // 
            // button_Stop
            // 
            this.button_Stop.Location = new System.Drawing.Point(551, 223);
            this.button_Stop.Name = "button_Stop";
            this.button_Stop.Size = new System.Drawing.Size(204, 54);
            this.button_Stop.TabIndex = 1;
            this.button_Stop.Text = "Stop";
            this.button_Stop.UseVisualStyleBackColor = true;
            this.button_Stop.Click += new System.EventHandler(this.button_Stop_Click);
            // 
            // textBox_Port
            // 
            this.textBox_Port.Location = new System.Drawing.Point(22, 59);
            this.textBox_Port.Name = "textBox_Port";
            this.textBox_Port.Size = new System.Drawing.Size(145, 31);
            this.textBox_Port.TabIndex = 2;
            this.toolTip1.SetToolTip(this.textBox_Port, "Ex:70890");
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(29, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 24);
            this.label1.TabIndex = 3;
            this.label1.Text = "Port";
            // 
            // textBox_Url
            // 
            this.textBox_Url.Location = new System.Drawing.Point(187, 59);
            this.textBox_Url.Name = "textBox_Url";
            this.textBox_Url.Size = new System.Drawing.Size(578, 31);
            this.textBox_Url.TabIndex = 4;
            this.toolTip1.SetToolTip(this.textBox_Url, "http or https");
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(195, 32);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 24);
            this.label2.TabIndex = 5;
            this.label2.Text = "URL";
            // 
            // textBox_Log
            // 
            this.textBox_Log.Location = new System.Drawing.Point(12, 283);
            this.textBox_Log.Multiline = true;
            this.textBox_Log.Name = "textBox_Log";
            this.textBox_Log.ReadOnly = true;
            this.textBox_Log.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox_Log.Size = new System.Drawing.Size(743, 407);
            this.textBox_Log.TabIndex = 17;
            this.toolTip1.SetToolTip(this.textBox_Log, "Ex:123,456");
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            // 
            // buttonDefaultURL
            // 
            this.buttonDefaultURL.Location = new System.Drawing.Point(317, 119);
            this.buttonDefaultURL.Margin = new System.Windows.Forms.Padding(6);
            this.buttonDefaultURL.Name = "buttonDefaultURL";
            this.buttonDefaultURL.Size = new System.Drawing.Size(203, 46);
            this.buttonDefaultURL.TabIndex = 19;
            this.buttonDefaultURL.Text = "Default URL";
            this.buttonDefaultURL.UseVisualStyleBackColor = true;
            this.buttonDefaultURL.Click += new System.EventHandler(this.buttonDefaultURL_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.radioButtonHttps);
            this.groupBox2.Controls.Add(this.radioButtonHttp);
            this.groupBox2.Location = new System.Drawing.Point(22, 107);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(271, 100);
            this.groupBox2.TabIndex = 18;
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
            // buttonAddCert
            // 
            this.buttonAddCert.Location = new System.Drawing.Point(576, 119);
            this.buttonAddCert.Name = "buttonAddCert";
            this.buttonAddCert.Size = new System.Drawing.Size(179, 43);
            this.buttonAddCert.TabIndex = 20;
            this.buttonAddCert.Text = "Add Cert";
            this.buttonAddCert.UseVisualStyleBackColor = true;
            this.buttonAddCert.Click += new System.EventHandler(this.buttonAddCert_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 708);
            this.Controls.Add(this.buttonAddCert);
            this.Controls.Add(this.buttonDefaultURL);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.textBox_Log);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox_Url);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox_Port);
            this.Controls.Add(this.button_Stop);
            this.Controls.Add(this.button_Start);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Tiny REST Server";
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button_Start;
        private System.Windows.Forms.Button button_Stop;
        private System.Windows.Forms.TextBox textBox_Port;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_Url;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.TextBox textBox_Log;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Button buttonDefaultURL;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton radioButtonHttps;
        private System.Windows.Forms.RadioButton radioButtonHttp;
        private System.Windows.Forms.Button buttonAddCert;
    }
}

