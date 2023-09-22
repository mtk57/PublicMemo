using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TabOrderHelper
{
    public partial class Form1 : Form
    {
        private TabOrderHelper _tabHelper = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _tabHelper = new TabOrderHelper(this);
        }


    }
}
