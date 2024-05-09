using GrepMapping.Common;

namespace GrepMapping
{
    partial class MainForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.tbGrepResultFilePath = new System.Windows.Forms.TextBox();
            this.btnRefGrepResultFile = new System.Windows.Forms.Button();
            this.cmbGrepResultType = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cbIsRegEx1 = new System.Windows.Forms.CheckBox();
            this.cbIsContentsLineCommentDelete = new System.Windows.Forms.CheckBox();
            this.label5 = new System.Windows.Forms.Label();
            this.tbMappingResultDirPath = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnRefMappingResultDir = new System.Windows.Forms.Button();
            this.btnMappingStart = new System.Windows.Forms.Button();
            this.tbContentsReplaceBeforeWord1 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tbContentsReplaceAfterWord1 = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tabControl2 = new System.Windows.Forms.TabControl();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.dgvTranscription = new System.Windows.Forms.DataGridView();
            this.ClmApplyOrder = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ClmEnable = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.ClmMappingFilePath = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ClmSrcStartRow = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ClmSrcKey = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ClmSrcCopy = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ClmDstStartRow = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ClmDstKey = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ClmDstCopy = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ClmMultiLine = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.btnLoad = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnGrepResultEditStart = new System.Windows.Forms.Button();
            this.cbIsDebugLog = new System.Windows.Forms.CheckBox();
            this.tabControl.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabControl2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTranscription)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(116, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "Grep Result File Path";
            // 
            // tbGrepResultFilePath
            // 
            this.tbGrepResultFilePath.AllowDrop = true;
            this.tbGrepResultFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbGrepResultFilePath.Location = new System.Drawing.Point(24, 36);
            this.tbGrepResultFilePath.Name = "tbGrepResultFilePath";
            this.tbGrepResultFilePath.Size = new System.Drawing.Size(703, 19);
            this.tbGrepResultFilePath.TabIndex = 1;
            this.tbGrepResultFilePath.TextChanged += new System.EventHandler(this.tbGrepResultFilePath_TextChanged);
            this.tbGrepResultFilePath.DragDrop += new System.Windows.Forms.DragEventHandler(this.tbGrepResultFilePath_DragDrop);
            this.tbGrepResultFilePath.DragEnter += new System.Windows.Forms.DragEventHandler(this.tbGrepResultFilePath_DragEnter);
            // 
            // btnRefGrepResultFile
            // 
            this.btnRefGrepResultFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefGrepResultFile.Location = new System.Drawing.Point(733, 33);
            this.btnRefGrepResultFile.Name = "btnRefGrepResultFile";
            this.btnRefGrepResultFile.Size = new System.Drawing.Size(55, 24);
            this.btnRefGrepResultFile.TabIndex = 2;
            this.btnRefGrepResultFile.Text = "ref";
            this.btnRefGrepResultFile.UseVisualStyleBackColor = true;
            this.btnRefGrepResultFile.Click += new System.EventHandler(this.btnRefGrepResultFile_Click);
            // 
            // cmbGrepResultType
            // 
            this.cmbGrepResultType.DisplayMember = "aa";
            this.cmbGrepResultType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbGrepResultType.FormattingEnabled = true;
            this.cmbGrepResultType.Items.AddRange(new object[] {
            "sakura"});
            this.cmbGrepResultType.Location = new System.Drawing.Point(17, 31);
            this.cmbGrepResultType.MaxDropDownItems = 1;
            this.cmbGrepResultType.Name = "cmbGrepResultType";
            this.cmbGrepResultType.Size = new System.Drawing.Size(121, 20);
            this.cmbGrepResultType.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(95, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "Grep Result Type";
            // 
            // cbIsRegEx1
            // 
            this.cbIsRegEx1.AutoSize = true;
            this.cbIsRegEx1.Location = new System.Drawing.Point(511, 85);
            this.cbIsRegEx1.Name = "cbIsRegEx1";
            this.cbIsRegEx1.Size = new System.Drawing.Size(57, 16);
            this.cbIsRegEx1.TabIndex = 15;
            this.cbIsRegEx1.Text = "RegEx";
            this.cbIsRegEx1.UseVisualStyleBackColor = true;
            // 
            // cbIsContentsLineCommentDelete
            // 
            this.cbIsContentsLineCommentDelete.AutoSize = true;
            this.cbIsContentsLineCommentDelete.Location = new System.Drawing.Point(163, 33);
            this.cbIsContentsLineCommentDelete.Name = "cbIsContentsLineCommentDelete";
            this.cbIsContentsLineCommentDelete.Size = new System.Drawing.Size(210, 16);
            this.cbIsContentsLineCommentDelete.TabIndex = 14;
            this.cbIsContentsLineCommentDelete.Text = "Contents Line Comment Row Delete";
            this.cbIsContentsLineCommentDelete.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(264, 68);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(155, 12);
            this.label5.TabIndex = 13;
            this.label5.Text = "Contents Replace After Word";
            // 
            // tbMappingResultDirPath
            // 
            this.tbMappingResultDirPath.AllowDrop = true;
            this.tbMappingResultDirPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbMappingResultDirPath.Location = new System.Drawing.Point(24, 375);
            this.tbMappingResultDirPath.Name = "tbMappingResultDirPath";
            this.tbMappingResultDirPath.Size = new System.Drawing.Size(703, 19);
            this.tbMappingResultDirPath.TabIndex = 7;
            this.tbMappingResultDirPath.DragDrop += new System.Windows.Forms.DragEventHandler(this.tbMappingResultDirPath_DragDrop);
            this.tbMappingResultDirPath.DragEnter += new System.Windows.Forms.DragEventHandler(this.tbMappingResultDirPath_DragEnter);
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(26, 359);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(130, 12);
            this.label3.TabIndex = 8;
            this.label3.Text = "Mapping Result Dir Path";
            // 
            // btnRefMappingResultDir
            // 
            this.btnRefMappingResultDir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefMappingResultDir.Location = new System.Drawing.Point(733, 375);
            this.btnRefMappingResultDir.Name = "btnRefMappingResultDir";
            this.btnRefMappingResultDir.Size = new System.Drawing.Size(55, 24);
            this.btnRefMappingResultDir.TabIndex = 9;
            this.btnRefMappingResultDir.Text = "ref";
            this.btnRefMappingResultDir.UseVisualStyleBackColor = true;
            this.btnRefMappingResultDir.Click += new System.EventHandler(this.btnRefMappingResultDir_Click);
            // 
            // btnMappingStart
            // 
            this.btnMappingStart.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnMappingStart.Enabled = false;
            this.btnMappingStart.Location = new System.Drawing.Point(414, 406);
            this.btnMappingStart.Name = "btnMappingStart";
            this.btnMappingStart.Size = new System.Drawing.Size(116, 27);
            this.btnMappingStart.TabIndex = 10;
            this.btnMappingStart.Text = "Mapping START";
            this.btnMappingStart.UseVisualStyleBackColor = true;
            this.btnMappingStart.Click += new System.EventHandler(this.btnMappingStart_Click);
            // 
            // tbContentsReplaceBeforeWord1
            // 
            this.tbContentsReplaceBeforeWord1.Location = new System.Drawing.Point(17, 83);
            this.tbContentsReplaceBeforeWord1.Name = "tbContentsReplaceBeforeWord1";
            this.tbContentsReplaceBeforeWord1.Size = new System.Drawing.Size(241, 19);
            this.tbContentsReplaceBeforeWord1.TabIndex = 11;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(15, 68);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(163, 12);
            this.label4.TabIndex = 11;
            this.label4.Text = "Contents Replace Before Word";
            // 
            // tabControl
            // 
            this.tabControl.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl.Controls.Add(this.tabPage1);
            this.tabControl.Controls.Add(this.tabPage2);
            this.tabControl.Location = new System.Drawing.Point(22, 67);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(764, 279);
            this.tabControl.TabIndex = 11;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.cbIsContentsLineCommentDelete);
            this.tabPage1.Controls.Add(this.cbIsRegEx1);
            this.tabPage1.Controls.Add(this.tbContentsReplaceAfterWord1);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.cmbGrepResultType);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.tbContentsReplaceBeforeWord1);
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(756, 253);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Grep Result Editing";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tbContentsReplaceAfterWord1
            // 
            this.tbContentsReplaceAfterWord1.Location = new System.Drawing.Point(264, 83);
            this.tbContentsReplaceAfterWord1.Name = "tbContentsReplaceAfterWord1";
            this.tbContentsReplaceAfterWord1.Size = new System.Drawing.Size(241, 19);
            this.tbContentsReplaceAfterWord1.TabIndex = 14;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.tabControl2);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(756, 253);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Mapping Rules";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tabControl2
            // 
            this.tabControl2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl2.Controls.Add(this.tabPage3);
            this.tabControl2.Location = new System.Drawing.Point(19, 20);
            this.tabControl2.Name = "tabControl2";
            this.tabControl2.SelectedIndex = 0;
            this.tabControl2.Size = new System.Drawing.Size(713, 210);
            this.tabControl2.TabIndex = 0;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dgvTranscription);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(705, 184);
            this.tabPage3.TabIndex = 0;
            this.tabPage3.Text = "Transcription";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // dgvTranscription
            // 
            this.dgvTranscription.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvTranscription.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dgvTranscription.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            this.dgvTranscription.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTranscription.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ClmApplyOrder,
            this.ClmEnable,
            this.ClmMappingFilePath,
            this.ClmSrcStartRow,
            this.ClmSrcKey,
            this.ClmSrcCopy,
            this.ClmDstStartRow,
            this.ClmDstKey,
            this.ClmDstCopy,
            this.ClmMultiLine});
            this.dgvTranscription.Location = new System.Drawing.Point(6, 17);
            this.dgvTranscription.Name = "dgvTranscription";
            this.dgvTranscription.RowTemplate.Height = 21;
            this.dgvTranscription.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvTranscription.Size = new System.Drawing.Size(693, 135);
            this.dgvTranscription.TabIndex = 0;
            // 
            // ClmApplyOrder
            // 
            this.ClmApplyOrder.HeaderText = "ApplyOrder";
            this.ClmApplyOrder.Name = "ClmApplyOrder";
            this.ClmApplyOrder.Width = 87;
            // 
            // ClmEnable
            // 
            this.ClmEnable.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.ClmEnable.HeaderText = "Enable";
            this.ClmEnable.Name = "ClmEnable";
            this.ClmEnable.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // ClmMappingFilePath
            // 
            this.ClmMappingFilePath.HeaderText = "MappingFilePath";
            this.ClmMappingFilePath.Name = "ClmMappingFilePath";
            this.ClmMappingFilePath.Width = 114;
            // 
            // ClmSrcStartRow
            // 
            this.ClmSrcStartRow.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.ClmSrcStartRow.HeaderText = "SrcStartRow";
            this.ClmSrcStartRow.Name = "ClmSrcStartRow";
            // 
            // ClmSrcKey
            // 
            this.ClmSrcKey.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.ClmSrcKey.HeaderText = "SrcKey";
            this.ClmSrcKey.Name = "ClmSrcKey";
            // 
            // ClmSrcCopy
            // 
            this.ClmSrcCopy.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.ClmSrcCopy.HeaderText = "SrcCopy";
            this.ClmSrcCopy.Name = "ClmSrcCopy";
            // 
            // ClmDstStartRow
            // 
            this.ClmDstStartRow.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.ClmDstStartRow.HeaderText = "DstStartRow";
            this.ClmDstStartRow.Name = "ClmDstStartRow";
            // 
            // ClmDstKey
            // 
            this.ClmDstKey.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.ClmDstKey.HeaderText = "DstKey";
            this.ClmDstKey.Name = "ClmDstKey";
            // 
            // ClmDstCopy
            // 
            this.ClmDstCopy.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.ClmDstCopy.HeaderText = "DstCopy";
            this.ClmDstCopy.Name = "ClmDstCopy";
            // 
            // ClmMultiLine
            // 
            this.ClmMultiLine.HeaderText = "MultiLine";
            this.ClmMultiLine.Name = "ClmMultiLine";
            this.ClmMultiLine.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ClmMultiLine.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.ClmMultiLine.Width = 76;
            // 
            // btnLoad
            // 
            this.btnLoad.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnLoad.Location = new System.Drawing.Point(24, 406);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(55, 36);
            this.btnLoad.TabIndex = 12;
            this.btnLoad.Text = "Config LOAD";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // btnSave
            // 
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSave.Location = new System.Drawing.Point(83, 406);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(55, 36);
            this.btnSave.TabIndex = 13;
            this.btnSave.Text = "Config SAVE";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnGrepResultEditStart
            // 
            this.btnGrepResultEditStart.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnGrepResultEditStart.Enabled = false;
            this.btnGrepResultEditStart.Location = new System.Drawing.Point(292, 406);
            this.btnGrepResultEditStart.Name = "btnGrepResultEditStart";
            this.btnGrepResultEditStart.Size = new System.Drawing.Size(116, 27);
            this.btnGrepResultEditStart.TabIndex = 14;
            this.btnGrepResultEditStart.Text = "Edit START";
            this.btnGrepResultEditStart.UseVisualStyleBackColor = true;
            this.btnGrepResultEditStart.Click += new System.EventHandler(this.btnGrepResultEditStart_Click);
            // 
            // cbIsDebugLog
            // 
            this.cbIsDebugLog.AutoSize = true;
            this.cbIsDebugLog.Location = new System.Drawing.Point(175, 412);
            this.cbIsDebugLog.Name = "cbIsDebugLog";
            this.cbIsDebugLog.Size = new System.Drawing.Size(78, 16);
            this.cbIsDebugLog.TabIndex = 15;
            this.cbIsDebugLog.Text = "Debug Log";
            this.cbIsDebugLog.UseVisualStyleBackColor = true;
            this.cbIsDebugLog.CheckedChanged += new System.EventHandler(this.cbIsDebugLog_CheckedChanged);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(801, 445);
            this.Controls.Add(this.cbIsDebugLog);
            this.Controls.Add(this.btnGrepResultEditStart);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnLoad);
            this.Controls.Add(this.tabControl);
            this.Controls.Add(this.btnMappingStart);
            this.Controls.Add(this.btnRefMappingResultDir);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.tbMappingResultDirPath);
            this.Controls.Add(this.btnRefGrepResultFile);
            this.Controls.Add(this.tbGrepResultFilePath);
            this.Controls.Add(this.label1);
            this.MinimumSize = new System.Drawing.Size(650, 435);
            this.Name = "MainForm";
            this.Text = "Simple Grep Mapping";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.tabControl.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabControl2.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvTranscription)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbGrepResultFilePath;
        private System.Windows.Forms.Button btnRefGrepResultFile;
        private System.Windows.Forms.ComboBox cmbGrepResultType;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbMappingResultDirPath;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnRefMappingResultDir;
        private System.Windows.Forms.Button btnMappingStart;
        private System.Windows.Forms.CheckBox cbIsContentsLineCommentDelete;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tbContentsReplaceBeforeWord1;
        private System.Windows.Forms.CheckBox cbIsRegEx1;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TextBox tbContentsReplaceAfterWord1;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TabControl tabControl2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.DataGridView dgvTranscription;
        private System.Windows.Forms.Button btnGrepResultEditStart;
        private System.Windows.Forms.DataGridViewTextBoxColumn ClmApplyOrder;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ClmEnable;
        private System.Windows.Forms.DataGridViewTextBoxColumn ClmMappingFilePath;
        private System.Windows.Forms.DataGridViewTextBoxColumn ClmSrcStartRow;
        private System.Windows.Forms.DataGridViewTextBoxColumn ClmSrcKey;
        private System.Windows.Forms.DataGridViewTextBoxColumn ClmSrcCopy;
        private System.Windows.Forms.DataGridViewTextBoxColumn ClmDstStartRow;
        private System.Windows.Forms.DataGridViewTextBoxColumn ClmDstKey;
        private System.Windows.Forms.DataGridViewTextBoxColumn ClmDstCopy;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ClmMultiLine;
        private System.Windows.Forms.CheckBox cbIsDebugLog;
    }
}

