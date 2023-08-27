using System.Windows.Forms;

namespace CustomTextBox
{
    public partial class Test00CustomTextBox : System.Windows.Forms.TextBox
    {
        public Test00CustomTextBox()
        {
            InitializeComponent();
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            base.OnPaint(pe);
        }
    }
}
