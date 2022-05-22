using NsComLibTest;
using System;
using System.Windows.Forms;

namespace ComLibTestKicker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonGetUserInfos_Click(object sender, EventArgs e)
        {
            try
            {
                var api = new ComLibMain();
                var userInfos = api.GetUserInfos();

                textBoxLog.Text = userInfos.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}\n{ex.StackTrace}");
            }
        }
    }
}
