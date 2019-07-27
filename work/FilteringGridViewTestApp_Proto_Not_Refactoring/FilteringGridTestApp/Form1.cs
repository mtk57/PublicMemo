using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace FilteringGridTestApp
{
    public partial class Form1 : Form
    {
        private DataTable _dbData = null;

        private BindingSource _dataSource = null;

        public Form1 ()
        {
            InitializeComponent();
        }

        #region Event handler
        private void Form1_Load ( object sender, EventArgs e )
        {
            init();
        }

        private void ButtonFiltering_Click ( object sender, EventArgs e )
        {
            doFiltering();
        }

        private void ButtonConfirm_Click ( object sender, EventArgs e )
        {
            confirm();
        }

        private void ButtonExit_Click ( object sender, EventArgs e )
        {
            this.Close();
        }

        private void Form1_FormClosing ( object sender, FormClosingEventArgs e )
        {
            if ( !isExit() )
            {
                e.Cancel = true;
                return;
            }

            // TBD  ここでlayoutIdListをapp.configに保存??
        }

        private void MenuItemImport_Click ( object sender, EventArgs e )
        {
            import();
        }

        private void MenuItemExport_Click ( object sender, EventArgs e )
        {
            export();
        }

        private void MenuItemAllClear_Click ( object sender, EventArgs e )
        {
            allClear();
        }
        #endregion Event handler

        #region Method
        private void init ()
        {
            initUI();

            autoSelectServiceCombo( "A0" );

            doFiltering();
        }

        private void initUI ()
        {
            _dbData = Utils.ConvertCsvToDataTable( @"..\..\DBData\DBData.tsv", '\t' );

            var clm = new DataColumn( LayoutInfo.NameIsDelete, typeof( bool ) );
            clm.DefaultValue = false;

            _dbData.Columns.Add( clm );

            bindServiceCombo();

            bindProductCombo();

            bindLayoutGrid();
        }

        private void bindServiceCombo ()
        {
            var table = _dbData.Copy();
            table.Columns.Remove( LayoutInfo.NameLayoutId );
            table.Columns.Remove( LayoutInfo.NameLayoutName );
            table.Columns.Remove( LayoutInfo.NameProductId );
            table.Columns.Remove( LayoutInfo.NameProductName );
            table.Columns.Remove( LayoutInfo.NameIsDelete );

            _comboBoxService.DataSource = table.Distinct( LayoutInfo.NameServiceId, LayoutInfo.NameServiceName );
            _comboBoxService.ValueMember = LayoutInfo.NameServiceId;
            _comboBoxService.DisplayMember = LayoutInfo.NameServiceName;
        }

        private void bindProductCombo ()
        {
            var table = _dbData.Copy();
            table.Columns.Remove( LayoutInfo.NameLayoutId );
            table.Columns.Remove( LayoutInfo.NameLayoutName );
            table.Columns.Remove( LayoutInfo.NameServiceId );
            table.Columns.Remove( LayoutInfo.NameServiceName );
            table.Columns.Remove( LayoutInfo.NameIsDelete );

            table.Rows.InsertAt( table.NewRow(), 0 );

            _comboBoxProduct.DataSource = table.Distinct( LayoutInfo.NameProductId, LayoutInfo.NameProductName );
            _comboBoxProduct.ValueMember = LayoutInfo.NameProductId;
            _comboBoxProduct.DisplayMember = LayoutInfo.NameProductName;
        }

        private void bindLayoutGrid ()
        {
            _dataSource = new BindingSource();
            _dataSource.DataSource = new DataView( _dbData );

            _dataGridView.AutoGenerateColumns = false;
            _dataGridView.DataSource = _dataSource;
        }

        private void autoSelectServiceCombo ( string serviceId )
        {
            _comboBoxService.SelectedValue = serviceId;
        }

        private void doFiltering ()
        {
            var cond1 = Utils.CreateFilterCondition( LayoutInfo.NameServiceId, _comboBoxService.SelectedValue.ToString() );
            var cond2 = Utils.CreateFilterCondition( LayoutInfo.NameProductId, _comboBoxProduct.SelectedValue.ToString() );

            var filter = Utils.JoinCondition( Utils.Separeator.and, cond1, cond2 );

            _dataSource.Filter = filter;

            rebindProductCombo();
        }

        private void rebindProductCombo ()
        {
            var table = getTableFromGrid();
            table.Columns.Remove( LayoutInfo.NameLayoutId );
            table.Columns.Remove( LayoutInfo.NameLayoutName );
            table.Columns.Remove( LayoutInfo.NameServiceId );
            table.Columns.Remove( LayoutInfo.NameServiceName );
            table.Columns.Remove( LayoutInfo.NameIsDelete );

            table.Rows.InsertAt( table.NewRow(), 0 );

            _comboBoxProduct.DataSource = table.Distinct( LayoutInfo.NameProductId, LayoutInfo.NameProductName );
        }

        private DataTable getTableFromGrid ()
            => ( ( _dataGridView.DataSource as BindingSource ).DataSource as DataView ).ToTable();

        private DataTable getCheckedTable ()
        {
            var view = new DataView( _dbData.Copy() );

            view.RowFilter = "IsDelete = true";

            return view.ToTable();
        }

        private void import ()
        {
            var path = Utils.SelectFileDialog( Utils.FILTER_ALL, "Import layoutId list?" );

            if ( string.IsNullOrEmpty( path ) ) return;

            allClear( isShowDialog: false );

            var list = Utils.ReadTextFile( path );

            var notApplyList = applyLayoutIdList( list );

            if ( notApplyList.Count > 0 )
            {
                Utils.ShowError( $"LayoutId apply failed! : \n{string.Join( "\n", notApplyList )}" );
            }
        }

        private void export ()
        {
            if ( !isExport )
            {
                Utils.ShowError("Checked row nothing!");
                return;
            }

            var path = Utils.SaveFileDialog( Utils.FILTER_ALL, "Export layoutId list?" );

            if ( string.IsNullOrEmpty( path ) ) return;

            var list = Utils.GetLayoutIdList( _dbData );

            Utils.WriteTextFile( path, list );
        }

        private void allClear (bool isShowDialog = true)
        {
            if ( isShowDialog &&
                 Utils.ShowWarning( "All clear?" ) == DialogResult.Cancel )
            {
                return;
            }

            foreach ( DataRow row in _dbData.Rows )
            {
                row [ LayoutInfo.NameIsDelete ] = false;
            }
        }

        private List<string> applyLayoutIdList ( List<string> importList )
        {
            var notApplyList = new List<string>();

            foreach ( var layoutId in importList )
            {
                if ( _dbData.AsEnumerable()
                        .Any( row => row[LayoutInfo.NameLayoutId].ToString() == layoutId ) )
                {
                    foreach ( var row in _dbData.AsEnumerable()
                        .Where( row => row [ LayoutInfo.NameLayoutId ].ToString() == layoutId ) )
                    {
                        row [ LayoutInfo.NameIsDelete ] = true;
                    }
                }
                else
                {
                    notApplyList.Add( layoutId );
                }
            }

            return notApplyList;
        }

        private bool isExport => getCheckedTable().Rows.Count > 0;

        private void confirm ()
        {
            var table = getCheckedTable();

            if ( table.Rows.Count <= 0 )
            {
                Utils.ShowError( "Checked row nothing!" );
                return;
            }

            using ( var dialog = new Form2( table ) )
            {
                dialog.ShowDialog();
            }
        }

        private bool isExit ()
        {
            if ( Utils.ShowWarning( "Are you exit?" ) == DialogResult.Cancel ) return false;
            return true;
        }
        #endregion Method
    }
}
