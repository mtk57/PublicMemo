using nsLibCOM;
using System;
using System.Windows.Forms;

namespace TinyHttpClient2
{
    public partial class Form1 : Form
    {
        private LibCOM _lib = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                _lib = new LibCOM();
                _lib.Init();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (_lib != null)
            {
                _lib.Dispose();
            }
        }

        private void buttonPostToken_Click(object sender, EventArgs e)
        {
            if (_lib == null)
            {
                MessageBox.Show("LibCOM is null!");
                return;
            }

            try
            {
                textBoxToken.Text = _lib.PostToken();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void buttonReload_Click(object sender, EventArgs e)
        {
            if (_lib == null)
            {
                MessageBox.Show("LibCOM is null!");
                return;
            }

            try
            {
                _lib.ReloadSettings();
                MessageBox.Show("Reload success.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }
    }
}
