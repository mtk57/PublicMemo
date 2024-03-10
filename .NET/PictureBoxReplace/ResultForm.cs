using System.Data;
using System.Windows.Forms;

namespace PictureBoxReplace
{
    public partial class ResultForm : Form
    {
        public ResultForm ( DataTable result)
        {
            InitializeComponent();

            if ( result == null )
            {
                dgvResult.Visible = false;

                lblResultFiles.Text = "0";

                return;
            }

            dgvResult.DataSource = result;

            dgvResult.ReadOnly = true;

            lblResultFiles.Text = result.Rows.Count.ToString();
        }


    }
}
