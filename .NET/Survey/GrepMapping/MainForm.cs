using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using GrepMapping.Common;
using GrepMapping.Config;
using GrepMapping.Proc;

namespace GrepMapping
{
    public partial class MainForm : Form
    {
        private Configs _config;
        private MappingDataTable _mapping;
        private MappingResultTable _result;
        private MappingResultForm _resultForm;

        public MainForm ()
        {
            InitializeComponent();

            Logger.Initialize();

            _config = new Configs();
            _mapping = new MappingDataTable();
            _result = new MappingResultTable();
        }

        #region Event handler
        private void MainForm_Load ( object sender, EventArgs e )
        {
            cmbGrepResultType.SelectedIndex = 0;
        }

        private void btnRefGrepResultFile_Click ( object sender, EventArgs e )
        {
            tbGrepResultFilePath.Text = Utils.OpenFileDialog( "Grep結果ファイルの選択", "テキストファイル(*.*)|*.*" );
        }

        private void tbGrepResultFilePath_DragDrop ( object sender, DragEventArgs e )
        {
            tbGrepResultFilePath.Text = Utils.GetDropData( e );
        }

        private void tbGrepResultFilePath_DragEnter ( object sender, DragEventArgs e )
        {
            Utils.StartDragDrop( e );
        }

        private void tbGrepResultFilePath_TextChanged ( object sender, EventArgs e )
        {
            if ( tbGrepResultFilePath.Text.Length == 0 )
            {
                btnGrepResultEditStart.Enabled = false;
                return;
            }
            btnGrepResultEditStart.Enabled = true;
        }

        private void btnRefMappingResultDir_Click ( object sender, EventArgs e )
        {
            tbMappingResultDirPath.Text = Utils.FolderBrowserDialog();
        }

        private void tbMappingResultDirPath_DragDrop ( object sender, DragEventArgs e )
        {
            tbMappingResultDirPath.Text = Utils.GetDropData( e );
        }

        private void tbMappingResultDirPath_DragEnter ( object sender, DragEventArgs e )
        {
            Utils.StartDragDrop( e );
        }

        private void btnGrepResultEditStart_Click ( object sender, EventArgs e )
        {
            try
            {
                if ( Utils.ShowOkCancelMessageBox() == DialogResult.Cancel )
                {
                    return;
                }

                btnMappingStart.Enabled = false;

                ConfigFromForm();

                _result = new MappingResultTable();

                var grepEdit = new GrepResultEdit( _config, _result );
                grepEdit.Run();

                _resultForm = new MappingResultForm( _result );
                _resultForm.Show();

                btnMappingStart.Enabled = true;

                //MessageBox.Show( "Succeess!" );
            }
            catch ( Exception ex )
            {
                Logger.Error( ex.ToString() );
                MessageBox.Show( ex.ToString() );
            }
        }

        private void btnMappingStart_Click ( object sender, EventArgs e )
        {
            try
            {
                if ( Utils.ShowOkCancelMessageBox() == DialogResult.Cancel )
                {
                    return;
                }

                ConfigFromForm();

                if ( string.IsNullOrEmpty( _config.MappingResultDirPath ) ||
                    !Directory.Exists( _config.MappingResultDirPath ) )
                {
                    throw new DirectoryNotFoundException( $"Directory is not exist! <{_config.MappingResultDirPath}>" );
                }

                _resultForm.Close();
                _resultForm.Dispose();
                _resultForm = null;

                var mapping = new Mapping(_config, _result);
                mapping.Run();

                var outputTsv = new Output( _config, _result );
                outputTsv.Run();

                _resultForm = new MappingResultForm( _result );
                _resultForm.Show();

                MessageBox.Show("Succeess!");
            }
            catch ( Exception ex )
            {
                Logger.Error(ex.ToString());
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnLoad_Click ( object sender, EventArgs e )
        {
            try
            {
                if ( Utils.ShowOkCancelMessageBox() == DialogResult.Cancel )
                {
                    return;
                }

                _config = ConfigUtil.ReadSettings();

                ConfigToForm();
            }
            catch ( Exception ex )
            {
                Logger.Error( ex.ToString() );
                MessageBox.Show( ex.ToString() );
            }
        }

        private void btnSave_Click ( object sender, EventArgs e )
        {
            try
            {
                if ( Utils.ShowOkCancelMessageBox() == DialogResult.Cancel )
                {
                    return;
                }

                ConfigFromForm();

                ConfigUtil.WriteSetting( _config );
            }
            catch ( Exception ex )
            {
                Logger.Error( ex.ToString() );
                MessageBox.Show( ex.ToString() );
            }
        }
        #endregion Event handler

        private void ConfigToForm ()
        {
            tbGrepResultFilePath.Text = _config.GrepResultFilePath;
            tbMappingResultDirPath.Text = _config.MappingResultDirPath;

            // GrepResultEditingInfo
            cmbGrepResultType.Text = _config.GrepResultEditing.GrepResultType;
            cbIsContentsLineCommentDelete.Checked = _config.GrepResultEditing.IsContentsLineCommentRowDelete;

            tbContentsReplaceBeforeWord1.Text = string.Empty;
            tbContentsReplaceAfterWord1.Text = string.Empty;
            cbIsRegEx1.Checked = false;

            if ( _config.GrepResultEditing.ReplaceWords != null && _config.GrepResultEditing.ReplaceWords.Count > 0 )
            {
                tbContentsReplaceBeforeWord1.Text = _config.GrepResultEditing.ReplaceWords [0].ContentsReplaceBeforeWord;
                tbContentsReplaceAfterWord1.Text = _config.GrepResultEditing.ReplaceWords [0].ContentsReplaceAfterWord;
                cbIsRegEx1.Checked = _config.GrepResultEditing.ReplaceWords [0].IsRegEx;
            }

            // MappingRules
            dgvTranscription.Columns.Clear();
            dgvTranscription.Rows.Clear();

            _mapping = new MappingDataTable( _config );

            dgvTranscription.DataSource = _mapping;
        }

        private void ConfigFromForm ()
        {
            _config.GrepResultFilePath = tbGrepResultFilePath.Text;
            _config.MappingResultDirPath = tbMappingResultDirPath.Text;

            // GrepResultEditingInfo
            var editInfo = new GrepResultEditingInfo();
            editInfo.GrepResultType = cmbGrepResultType.Text;
            editInfo.IsContentsLineCommentRowDelete = cbIsContentsLineCommentDelete.Checked;

            var replaceInfo1 = new ReplaceWordInfo();
            replaceInfo1.ContentsReplaceBeforeWord = tbContentsReplaceBeforeWord1.Text;
            replaceInfo1.ContentsReplaceAfterWord = tbContentsReplaceAfterWord1.Text;
            replaceInfo1.IsRegEx = cbIsRegEx1.Checked;

            editInfo.ReplaceWords = new List<ReplaceWordInfo>();
            editInfo.ReplaceWords.Add( replaceInfo1 );

            _config.GrepResultEditing = editInfo;

            // MappingRules
            _config.MappingRules = new List<MappingRuleInfo>();

            var data = ( MappingDataTable ) dgvTranscription.DataSource;

            if ( data != null && data.Validate() )
            {
                foreach ( DataRow row in data.Rows )
                {
                    var rule = new MappingRuleInfo();

                    rule.ApplyOrder = int.Parse( row [ Const.CLM_APPLY_ORDER ].ToString() );
                    rule.IsEnable = bool.Parse( row [ Const.CLM_ENABLE ].ToString() );
                    rule.MappingFilePath = row [ Const.CLM_MAPPING_FILE_PATH ].ToString();
                    rule.SrcStartRow = int.Parse( row [ Const.CLM_SRC_START_ROW ].ToString() );
                    rule.SrcKey = int.Parse( row [ Const.CLM_SRC_KEY ].ToString() );
                    rule.SrcCopy = int.Parse( row [ Const.CLM_SRC_COPY ].ToString() );
                    rule.DstStartRow = int.Parse( row [ Const.CLM_DST_START_ROW ].ToString() );
                    rule.DstKey = int.Parse( row [ Const.CLM_DST_KEY ].ToString() );
                    rule.DstCopy = int.Parse( row [ Const.CLM_DST_COPY ].ToString() );

                    _config.MappingRules.Add( rule );
                }
            }
        }


    }
}
