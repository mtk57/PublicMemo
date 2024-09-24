using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace BarcodeReceiver
{
    public partial class Form1 : Form
    {
        public Form1 ()
        {
            InitializeComponent();
        }

        private void Form1_KeyPress ( object sender, KeyPressEventArgs e )
        {
            Debug.WriteLine( $"{e.KeyChar}" );

            textBoxBarcodeData.AppendText( e.KeyChar.ToString() );

            if ( e.KeyChar == ( char ) Keys.Enter )
            {
                ProcessBarcode( textBoxBarcodeData.Text );
                textBoxBarcodeData.Clear();
                e.Handled = true;
            }
        }

        private void ProcessBarcode ( string barcode )
        {
            if ( !string.IsNullOrWhiteSpace( barcode ) )
            {
                listBoxBarcodes.Items.Add( $"{DateTime.Now:HH:mm:ss} - {barcode}" );
                listBoxBarcodes.SelectedIndex = listBoxBarcodes.Items.Count - 1;
            }
        }
    }
}
