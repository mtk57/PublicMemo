namespace FilteringGridTestApp
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose ( bool disposing )
        {
            if ( disposing && ( components != null ) )
            {
                components.Dispose();
            }
            base.Dispose( disposing );
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent ()
        {
            this._buttonCancel = new System.Windows.Forms.Button();
            this._buttonRun = new System.Windows.Forms.Button();
            this._dataGridView = new System.Windows.Forms.DataGridView();
            this.ColumnLayoutId = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnLayoutName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnServiceId = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnServiceName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnProductId = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnProductName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnIsDelete = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this._labelCount = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this._dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // _buttonCancel
            // 
            this._buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this._buttonCancel.Location = new System.Drawing.Point(26, 198);
            this._buttonCancel.Name = "_buttonCancel";
            this._buttonCancel.Size = new System.Drawing.Size(75, 23);
            this._buttonCancel.TabIndex = 5;
            this._buttonCancel.Text = "Cancel";
            this._buttonCancel.UseVisualStyleBackColor = true;
            // 
            // _buttonRun
            // 
            this._buttonRun.DialogResult = System.Windows.Forms.DialogResult.OK;
            this._buttonRun.Location = new System.Drawing.Point(416, 198);
            this._buttonRun.Name = "_buttonRun";
            this._buttonRun.Size = new System.Drawing.Size(75, 23);
            this._buttonRun.TabIndex = 6;
            this._buttonRun.Text = "Run";
            this._buttonRun.UseVisualStyleBackColor = true;
            this._buttonRun.Click += new System.EventHandler(this.ButtonRun_Click);
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
            this._dataGridView.Location = new System.Drawing.Point(26, 30);
            this._dataGridView.MultiSelect = false;
            this._dataGridView.Name = "_dataGridView";
            this._dataGridView.RowHeadersVisible = false;
            this._dataGridView.RowTemplate.Height = 21;
            this._dataGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this._dataGridView.Size = new System.Drawing.Size(465, 150);
            this._dataGridView.TabIndex = 7;
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
            // _labelCount
            // 
            this._labelCount.AutoSize = true;
            this._labelCount.Location = new System.Drawing.Point(351, 203);
            this._labelCount.Name = "_labelCount";
            this._labelCount.Size = new System.Drawing.Size(35, 12);
            this._labelCount.TabIndex = 8;
            this._labelCount.Text = "label1";
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(527, 250);
            this.Controls.Add(this._labelCount);
            this.Controls.Add(this._dataGridView);
            this.Controls.Add(this._buttonRun);
            this.Controls.Add(this._buttonCancel);
            this.Name = "Form2";
            this.Text = "Form2";
            ((System.ComponentModel.ISupportInitialize)(this._dataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button _buttonCancel;
        private System.Windows.Forms.Button _buttonRun;
        private System.Windows.Forms.DataGridView _dataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnLayoutId;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnLayoutName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnServiceId;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnServiceName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnProductId;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnProductName;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ColumnIsDelete;
        private System.Windows.Forms.Label _labelCount;
    }
}