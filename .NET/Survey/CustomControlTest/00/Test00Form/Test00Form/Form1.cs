using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Test00Form
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        protected override bool ProcessTabKey(bool forward)
        {
            //Console.WriteLine("ProcessTabKey");

            foreach (Control c in this.Controls)
            {
                if (c.Focused)
                {
                    Console.WriteLine($"Focused Control={c.Name}, TabIdx={c.TabIndex}, TabStop={c.TabStop}");
                    break;
                }
            }

            return base.ProcessTabKey(forward);
        }



    }
}
