using System.Windows.Forms;
using System.ComponentModel;

namespace TestUserControl
{
    public partial class TestUserTextBoxControl : UserControl
    {


        public TestUserTextBoxControl()
        {
            InitializeComponent();
        }

        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        [DefaultValue("")]
        public override string Text
        {
            get { return this.test00TextBoxEx1.Text; }
            set { this.test00TextBoxEx1.Text = value; }
        }

    }
}
