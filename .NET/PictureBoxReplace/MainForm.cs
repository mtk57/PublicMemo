using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace PictureBoxReplace
{
    public partial class MainForm : Form
    {
        private readonly DataModel _model = null;

        private DataTable _dtResult = null;

        public MainForm ()
        {
            InitializeComponent();

            _model = new DataModel();
        }

        #region Event handler
        private void BtnInResxPathRef_Click ( object sender, EventArgs e )
        {
            tbInResxPath.Text = Common.OpenFileDialog("resxファイルの選択", "resxファイル(*.resx)|*.resx" );

            UpdatePictureBoxNameComboBox();
        }

        private void BtnOutReplaceImageFilePathRef_Click ( object sender, EventArgs e )
        {
            tbOutReplaceImageFilePath.Text = Common.OpenFileDialog( "画像ファイルの選択", "画像ファイル|*.png;*.jpg;*.jpeg;*.bmp" );

            DrawReplaceImage();
        }

        private void BtnOutTargetDirPathRef_Click ( object sender, EventArgs e )
        {
            tbOutTargetDirPath.Text = Common.FolderBrowserDialog();

            CountResx();
        }

        private void BtnConfirmResult_Click ( object sender, EventArgs e )
        {
            var result = new ResultForm( _dtResult );

            result.ShowDialog();
        }

        private void TbInResxPath_DragDrop ( object sender, DragEventArgs e )
        {
            tbInResxPath.Text = Common.GetDropData( e );

            UpdatePictureBoxNameComboBox();
        }

        private void TbInResxPath_DragEnter ( object sender, DragEventArgs e )
        {
            Common.StartDragDrop( e );
        }

        private void TbOutReplaceImageFilePath_DragDrop ( object sender, DragEventArgs e )
        {
            tbOutReplaceImageFilePath.Text = Common.GetDropData( e );

            DrawReplaceImage();
        }

        private void TbOutReplaceImageFilePath_DragEnter ( object sender, DragEventArgs e )
        {
            Common.StartDragDrop( e );
        }

        private void TbOutTargetDirPath_DragDrop ( object sender, DragEventArgs e )
        {
            tbOutTargetDirPath.Text = Common.GetDropData( e );

            CountResx();
        }

        private void TbOutTargetDirPath_DragEnter ( object sender, DragEventArgs e )
        {
            Common.StartDragDrop( e );
        }

        private void TbInResxPath_Leave ( object sender, EventArgs e )
        {
            UpdatePictureBoxNameComboBox();
        }

        private void CmbInPictureBoxName_Leave ( object sender, EventArgs e )
        {
            DrawBeforeImage();
        }

        private void TbOutReplaceImageFilePath_Leave ( object sender, EventArgs e )
        {
            DrawReplaceImage();
        }

        private void TbOutTargetDirPath_Leave ( object sender, EventArgs e )
        {
            CountResx();
        }

        private void CmbInPictureBoxName_SelectedIndexChanged ( object sender, EventArgs e )
        {
            DrawBeforeImage();
        }

        private void BtnReplace_Click ( object sender, EventArgs e )
        {
            if ( Common.ShowOkCancelMessageBox() == DialogResult.Cancel ) return;

            _model.ResxPath = tbInResxPath.Text;
            _model.PictureBoxName = cmbInPictureBoxName.Text;
            _model.ImageData = tbInImageData.Text;
            _model.ReplaceImagePath = tbOutReplaceImageFilePath.Text;
            _model.TargetDirPath = tbOutTargetDirPath.Text;
            _model.IsBackup = cbOutCreateResxBak.Checked;
            _model.ReplaceImageData = "";

            ChangeControlEnabled( false );

            try
            {
                _dtResult = null;

                Replace.Validate( _model );

                Replace.Prepare();

                _dtResult = Replace.Execute();

                MessageBox.Show($"Success! ({_dtResult.Rows.Count} datas replaced.)");
            }
            catch ( Exception ex )
            {
                MessageBox.Show( $"Error! \r\n[Msg={ex.Message}]\n\n[Stack={ex.StackTrace}]" );
            }
            finally
            {
                ChangeControlEnabled( true );
            }
        }

        #endregion Event handler

        private void ChangeControlEnabled ( bool isEnabled )
        {
            gbBefore.Enabled = isEnabled;
            gbAfter.Enabled = isEnabled;

            btnReplace.Enabled = isEnabled;
            btnConfirmResult.Enabled = isEnabled;
        }

        private void UpdatePictureBoxNameComboBox ()
        {
            lblPictureBoxCount.Text = "";

            if ( !File.Exists( tbInResxPath.Text ) ) return;

            cmbInPictureBoxName.Items.Clear();

            var resxPath = tbInResxPath.Text;

            var pics = Resx.GetPictureBoxNames( resxPath );

            if ( pics.Count == 0 ) return;

            cmbInPictureBoxName.Items.AddRange( pics.ToArray() );

            cmbInPictureBoxName.SelectedIndex = 0;

            lblPictureBoxCount.Text = $"{pics.Count} Found.";
        }

        private void DrawBeforeImage ()
        {
            if ( string.IsNullOrEmpty( cmbInPictureBoxName.Text ) ) return;

            var picName = cmbInPictureBoxName.Text;

            if ( !File.Exists( tbInResxPath.Text ) ) return;

            var resxPath = tbInResxPath.Text;

            var imgData = Resx.GetImageData( resxPath, picName );

            if ( imgData == "" ) return;

            pbSelectImage.Image = Common.ConvertBase64StringToImage( imgData );

            tbInImageData.Text = imgData;
        }

        private void DrawReplaceImage ()
        {
            var path = tbOutReplaceImageFilePath.Text;

            if ( string.IsNullOrEmpty( path ) ) return;

            if ( !File.Exists( path ) ) return;

            pbReplaceImage.Image = Image.FromFile( path );
        }

        private void CountResx ()
        {
            lblResxCount.Text = "";

            var path = tbOutTargetDirPath.Text;

            if ( string.IsNullOrEmpty( path ) ) return;

            if ( !Directory.Exists( path ) ) return;
;
            if ( !Common.IsDir( path ) ) return;

            var resxPath = Common.GetResxFiles( path );

            lblResxCount.Text = $"{resxPath.Count()} Found.";
        }
    }
}
