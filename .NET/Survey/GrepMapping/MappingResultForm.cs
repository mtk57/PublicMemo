using System.Windows.Forms;
using GrepMapping.Proc;

namespace GrepMapping
{
    internal partial class MappingResultForm : Form
    {
        private MappingResultTable _result;

        private MappingResultForm ()
        {
        }

        public MappingResultForm ( MappingResultTable result )
        {
            InitializeComponent();

            _result = result;

            this.dgvMappingResult.DataSource = result;
        }
    }
}
