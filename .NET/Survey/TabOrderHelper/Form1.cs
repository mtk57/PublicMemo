using System;
using System.Windows.Forms;

namespace TabOrderHelper
{
    public partial class Form1 : Form
    {
        private TabOrderHelper _helper = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _helper = new TabOrderHelper(this);
        }

        //protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        //{
        //    var activeControl = this.ActiveControl;

        //    Console.WriteLine($"activeControl={activeControl.Name}");

        //    if (keyData == Keys.Tab)
        //    {
        //        // TABキーが押されたときの処理
        //        //System.Diagnostics.Debug.WriteLine("TAB key pressed");

        //        var nextControl = _helper.GetNextControl(activeControl);

        //        nextControl.Focus();

        //        return true; // イベントを処理済みとしてマークする
        //    }
        //    else if (keyData == (Keys.Shift | Keys.Tab))
        //    {
        //        // SHIFT+TABキーが押されたときの処理
        //        //System.Diagnostics.Debug.WriteLine("SHIFT+TAB key pressed");

        //        var prevControl = _helper.GetNextControl(activeControl, false);

        //        prevControl.Focus();

        //        return true; // イベントを処理済みとしてマークする
        //    }

        //    return base.ProcessCmdKey(ref msg, keyData);
        //}
    }
}
