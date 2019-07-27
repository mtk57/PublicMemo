namespace FilteringGridTestApp
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
            this._comboBoxService = new System.Windows.Forms.ComboBox();
            this._comboBoxProduct = new System.Windows.Forms.ComboBox();
            this._buttonFiltering = new System.Windows.Forms.Button();
            this._dataGridView = new System.Windows.Forms.DataGridView();
            this.ColumnLayoutId = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnLayoutName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnServiceId = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnServiceName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnProductId = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnProductName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnIsDelete = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this._buttonConfirm = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.allClearToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this._buttonExit = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this._dataGridView)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // _comboBoxService
            // 
            this._comboBoxService.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this._comboBoxService.FormattingEnabled = true;
            this._comboBoxService.Location = new System.Drawing.Point(27, 41);
            this._comboBoxService.Name = "_comboBoxService";
            this._comboBoxService.Size = new System.Drawing.Size(121, 20);
            this._comboBoxService.TabIndex = 0;
            // 
            // _comboBoxProduct
            // 
            this._comboBoxProduct.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this._comboBoxProduct.FormattingEnabled = true;
            this._comboBoxProduct.Location = new System.Drawing.Point(230, 41);
            this._comboBoxProduct.Name = "_comboBoxProduct";
            this._comboBoxProduct.Size = new System.Drawing.Size(121, 20);
            this._comboBoxProduct.TabIndex = 1;
            // 
            // _buttonFiltering
            // 
            this._buttonFiltering.Location = new System.Drawing.Point(417, 38);
            this._buttonFiltering.Name = "_buttonFiltering";
            this._buttonFiltering.Size = new System.Drawing.Size(75, 23);
            this._buttonFiltering.TabIndex = 2;
            this._buttonFiltering.Text = "Filtering";
            this._buttonFiltering.UseVisualStyleBackColor = true;
            this._buttonFiltering.Click += new System.EventHandler(this.ButtonFiltering_Click);
            // 
            // _dataGridView
            // 
            this._dataGridView.AllowUserToAddRows = false;
            this._dataGridView.AllowUserToDeleteRows = false;
            this._dataGridView.AllowUserToResizeRows = false;
            this._dataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this._dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this._dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColumnLayoutId,
            this.ColumnLayoutName,
            this.ColumnServiceId,
            this.ColumnServiceName,
            this.ColumnProductId,
            this.ColumnProductName,
            this.ColumnIsDelete});
            this._dataGridView.Location = new System.Drawing.Point(27, 87);
            this._dataGridView.MultiSelect = false;
            this._dataGridView.Name = "_dataGridView";
            this._dataGridView.RowHeadersVisible = false;
            this._dataGridView.RowTemplate.Height = 21;
            this._dataGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this._dataGridView.Size = new System.Drawing.Size(465, 150);
            this._dataGridView.TabIndex = 3;
            // 
            // ColumnLayoutId
            // 
            this.ColumnLayoutId.DataPropertyName = "LayoutId";
            this.ColumnLayoutId.HeaderText = "LayoutId";
            this.ColumnLayoutId.Name = "ColumnLayoutId";
            this.ColumnLayoutId.ReadOnly = true;
            this.ColumnLayoutId.Visible = false;
            // 
            // ColumnLayoutName
            // 
            this.ColumnLayoutName.DataPropertyName = "LayoutName";
            this.ColumnLayoutName.HeaderText = "LayoutName";
            this.ColumnLayoutName.Name = "ColumnLayoutName";
            this.ColumnLayoutName.ReadOnly = true;
            // 
            // ColumnServiceId
            // 
            this.ColumnServiceId.DataPropertyName = "ServiceId";
            this.ColumnServiceId.HeaderText = "ServiceId";
            this.ColumnServiceId.Name = "ColumnServiceId";
            this.ColumnServiceId.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColumnServiceId.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.ColumnServiceId.Visible = false;
            // 
            // ColumnServiceName
            // 
            this.ColumnServiceName.DataPropertyName = "ServiceName";
            this.ColumnServiceName.HeaderText = "ServiceName";
            this.ColumnServiceName.Name = "ColumnServiceName";
            // 
            // ColumnProductId
            // 
            this.ColumnProductId.DataPropertyName = "ProductId";
            this.ColumnProductId.HeaderText = "ProductId";
            this.ColumnProductId.Name = "ColumnProductId";
            this.ColumnProductId.Visible = false;
            // 
            // ColumnProductName
            // 
            this.ColumnProductName.DataPropertyName = "ProductName";
            this.ColumnProductName.HeaderText = "ProductName";
            this.ColumnProductName.Name = "ColumnProductName";
            // 
            // ColumnIsDelete
            // 
            this.ColumnIsDelete.DataPropertyName = "IsDelete";
            this.ColumnIsDelete.HeaderText = "IsDelete";
            this.ColumnIsDelete.Name = "ColumnIsDelete";
            this.ColumnIsDelete.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColumnIsDelete.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // _buttonConfirm
            // 
            this._buttonConfirm.Location = new System.Drawing.Point(417, 276);
            this._buttonConfirm.Name = "_buttonConfirm";
            this._buttonConfirm.Size = new System.Drawing.Size(75, 23);
            this._buttonConfirm.TabIndex = 5;
            this._buttonConfirm.Text = "Confirm";
            this._buttonConfirm.UseVisualStyleBackColor = true;
            this._buttonConfirm.Click += new System.EventHandler(this.ButtonConfirm_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.editToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(521, 24);
            this.menuStrip1.TabIndex = 6;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.importToolStripMenuItem,
            this.exportToolStripMenuItem});
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.openToolStripMenuItem.Text = "File";
            // 
            // importToolStripMenuItem
            // 
            this.importToolStripMenuItem.Name = "importToolStripMenuItem";
            this.importToolStripMenuItem.Size = new System.Drawing.Size(109, 22);
            this.importToolStripMenuItem.Text = "Import";
            this.importToolStripMenuItem.Click += new System.EventHandler(this.MenuItemImport_Click);
            // 
            // exportToolStripMenuItem
            // 
            this.exportToolStripMenuItem.Name = "exportToolStripMenuItem";
            this.exportToolStripMenuItem.Size = new System.Drawing.Size(109, 22);
            this.exportToolStripMenuItem.Text = "Export";
            this.exportToolStripMenuItem.Click += new System.EventHandler(this.MenuItemExport_Click);
            // 
            // editToolStripMenuItem
            // 
            this.editToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.allClearToolStripMenuItem});
            this.editToolStripMenuItem.Name = "editToolStripMenuItem";
            this.editToolStripMenuItem.Size = new System.Drawing.Size(39, 20);
            this.editToolStripMenuItem.Text = "Edit";
            // 
            // allClearToolStripMenuItem
            // 
            this.allClearToolStripMenuItem.Name = "allClearToolStripMenuItem";
            this.allClearToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.allClearToolStripMenuItem.Text = "AllClear";
            this.allClearToolStripMenuItem.Click += new System.EventHandler(this.MenuItemAllClear_Click);
            // 
            // _buttonExit
            // 
            this._buttonExit.Location = new System.Drawing.Point(27, 276);
            this._buttonExit.Name = "_buttonExit";
            this._buttonExit.Size = new System.Drawing.Size(75, 23);
            this._buttonExit.TabIndex = 7;
            this._buttonExit.Text = "Exit";
            this._buttonExit.UseVisualStyleBackColor = true;
            this._buttonExit.Click += new System.EventHandler(this.ButtonExit_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(521, 320);
            this.Controls.Add(this._buttonExit);
            this.Controls.Add(this._buttonConfirm);
            this.Controls.Add(this._dataGridView);
            this.Controls.Add(this._buttonFiltering);
            this.Controls.Add(this._comboBoxProduct);
            this.Controls.Add(this._comboBoxService);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Form1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this._dataGridView)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox _comboBoxService;
        private System.Windows.Forms.ComboBox _comboBoxProduct;
        private System.Windows.Forms.Button _buttonFiltering;
        private System.Windows.Forms.DataGridView _dataGridView;
        private System.Windows.Forms.Button _buttonConfirm;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exportToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem allClearToolStripMenuItem;
        private System.Windows.Forms.Button _buttonExit;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnLayoutId;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnLayoutName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnServiceId;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnServiceName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnProductId;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnProductName;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ColumnIsDelete;
    }
}

