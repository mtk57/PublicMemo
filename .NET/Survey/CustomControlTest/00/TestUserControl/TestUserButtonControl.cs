using System.Windows.Forms;
using System.ComponentModel;

namespace TestUserControl
{
    public partial class TestUserButtonControl : UserControl
    {
        public TestUserButtonControl()
        {
            InitializeComponent();
        }

        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        public override string Text
        {
            get { return this.test00ButtonEx1.Text; }
            set
            {
                this.test00ButtonEx1.UseMnemonic = true;
                this.test00ButtonEx1.Text = value;
                //this.test00ButtonEx1.Refresh();
                //this.Refresh();
            }
        }

        public override  bool Focused
        {
            get { return (this.test00ButtonEx1.Focused);
            }
        }
    }
}
