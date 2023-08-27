using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CustomTextBox
{
    public partial class Test00CustomButton : System.Windows.Forms.Button
    {
        public Test00CustomButton()
        {
            InitializeComponent();

            //this.Text = "Test"; 
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            base.OnPaint(pe);
        }
    }
}
