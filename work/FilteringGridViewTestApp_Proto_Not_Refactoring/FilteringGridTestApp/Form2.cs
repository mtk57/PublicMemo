using System;
using System.Data;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using System.Collections.Generic;

namespace FilteringGridTestApp
{
    public partial class Form2 : Form
    {
        private const string SQL_TEMPLATE = @"
begin;

UPDATE tableA
SET delete_flag = false
WHERE layout_id IN (
{0}
);

UPDATE tableB
SET dsp_flag = false
WHERE layout_id IN (
{0}
);

commit;
";

        private DataTable _dbData = null;

        public Form2 ( DataTable table )
        {
            InitializeComponent();

            _dbData = table;

            _labelCount.Text = string.Format( $"checked={_dbData.Rows.Count}" );

            _dataGridView.AutoGenerateColumns = false;
            _dataGridView.DataSource = _dbData;
            _dataGridView.Columns [ LayoutInfo.ColumnIsDelete ].Visible = false;
        }

        private void ButtonRun_Click ( object sender, EventArgs e )
        {
            run();
        }

        private void run ()
        {
            var path = Utils.SaveFileDialog( Utils.FILTER_SQL, "Create SQL?", fileName: "DeleteDesign.sql" );

            if ( string.IsNullOrEmpty( path ) ) return;

            var layoudIdList = Utils.GetLayoutIdList( _dbData );

            var list = layoudIdList.Select( x => string.Format( $"'{x}'" ) );

            var sb = new StringBuilder();

            sb.AppendLine( string.Format( SQL_TEMPLATE, string.Join( ",", list ) ) );

            var sql = new List<string>();

            sql.Add( sb.ToString() );

            Utils.WriteTextFile( path, sql, false, new UTF8Encoding( false ) );
        }
    }
}
