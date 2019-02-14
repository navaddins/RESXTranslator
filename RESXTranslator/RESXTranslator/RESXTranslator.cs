using System;
using System.Collections.Generic;
using System.Windows.Forms;
namespace RESXTranslator
{
    public partial class frmRESXTranslator : Form
    {
        private bool _FromGridClicked = true;
        private bool _AbortRequested;

        private string _DirSearch = Properties.Resources.DirSearch;
        private string _DirUnSearch = Properties.Resources.DirUnSearch;
        private string _DefaultFileDlgTitle = Properties.Resources.DefaultFileDlgTitle;
        private string _DefaultFileName = Properties.Resources.DefaultFileName;
        private string _DefaultExtension = Properties.Resources.DefaultExtension;
        private string _DefaultFilter = Properties.Resources.DefaultFilter;
        private string _StatusLineText;

        private Timer startupTimer;

        public frmRESXTranslator()
        {
            InitializeComponent();
        }

        private void frmRESXTranslator_Load(object sender, EventArgs e)
        {
            InitLoading();
        }

        private void btnSourceFileDlg_Click(object sender, EventArgs e)
        {
            FileFolderHelper _FileFolderHelper = new FileFolderHelper();
            string _PreviousFile = txtSourceRESXFile.Text;
            if (chkDir.Checked)
                txtSourceRESXFile.Text = _FileFolderHelper.FolderBrowserDialog(_PreviousFile, _DefaultFileDlgTitle);
            else
                txtSourceRESXFile.Text = _FileFolderHelper.OpenFileDialog(_PreviousFile, _DefaultFileDlgTitle, txtSearchPattern.Text, _DefaultExtension, _DefaultFilter);
            if (!_PreviousFile.ToLower().Equals(txtSourceRESXFile.Text.ToLower()))
                ClearGrid();
        }

        private void btnLoadRESX_Click(object sender, EventArgs e)
        {
            LoadRESXFileToGrid();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveResourceFile();
        }

        private void btnTranslate_Click(object sender, EventArgs e)
        {
            TranslateLanguage();
        }

        private void btnSaveSettings_Click(object sender, EventArgs e)
        {
            SaveTranslateAPISettings();
            SaveBingSpellCheckSettings();
            SaveProxySettings();
            DataGridViewCell _dgvc;
        }

        private void chkDirSearch_CheckedChanged(object sender, EventArgs e)
        {
            CheckDirSearch();
        }

        private void txtSearchPattern_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtSearchPattern.Text))
                SetDefaultFilter();
        }

        private void txtBingSpellCheckSubscriptionKey_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtBingSpellCheckSubscriptionKey.Text))
                chkBingSpellCheck.Checked = false;
            chkBingSpellCheck.Enabled = (!string.IsNullOrWhiteSpace(txtBingSpellCheckSubscriptionKey.Text));
        }

        private void cboSourceFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboSourceFilter.SelectedIndex != cboFromLanguage.SelectedIndex)
                cboFromLanguage.SelectedIndex = (cboFromLanguage.Items.Count <= 0) ? -1 : cboSourceFilter.SelectedIndex;
        }

        private void cboFromLanguage_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtSourceRESXFile.Clear();
            SetDefaultFilter();
            ClearGrid();
        }

        private void cboToLanguage_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void dgvFrom_Scroll(object sender, ScrollEventArgs e)
        {
            if ((dgvTo.Rows.Count > 0) && ((dgvTo.Rows.Count - 1) >= dgvFrom.FirstDisplayedScrollingRowIndex))
                dgvTo.FirstDisplayedScrollingRowIndex = dgvFrom.FirstDisplayedScrollingRowIndex;
        }

        private void dgvTo_Scroll(object sender, ScrollEventArgs e)
        {
            if (dgvTo.Rows.Count > 0)
                dgvFrom.FirstDisplayedScrollingRowIndex = dgvTo.FirstDisplayedScrollingRowIndex;
        }

        private void dgvFrom_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            if ((e.StateChanged != DataGridViewElementStates.Selected) ||
                (!_FromGridClicked) ||
                (dgvTo.Rows.Count <= 0) ||
                ((dgvTo.Rows.Count - 1) < e.Row.Index))
                return;

            dgvTo.ClearSelection();
            dgvTo.Rows[e.Row.Index].Selected = true;
        }

        private void dgvTo_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            if ((e.StateChanged != DataGridViewElementStates.Selected) ||
                (_FromGridClicked) ||
                (dgvFrom.Rows.Count <= 0) ||
                (dgvTo.Rows.Count <= 0))
                return;

            dgvFrom.ClearSelection();
            dgvFrom.Rows[e.Row.Index].Selected = true;
        }

        private void dgvFrom_Click(object sender, EventArgs e)
        {
            _FromGridClicked = true;
            if ((!_FromGridClicked) || (dgvTo.Rows.Count <= 0))
                return;

            DataGridViewSelectedRowCollection _DataGridViewSelectedRowCollection = ((DataGridView)sender).SelectedRows;
            DataGridViewRow _DataGridViewRow = (DataGridViewRow)_DataGridViewSelectedRowCollection[0];

            if ((dgvTo.Rows.Count - 1) < _DataGridViewRow.Index)
                return;

            dgvTo.ClearSelection();
            dgvTo.Rows[_DataGridViewRow.Index].Selected = true;
        }

        private void dgvTo_Click(object sender, EventArgs e)
        {
            _FromGridClicked = false;
            if ((_FromGridClicked) || (dgvTo.Rows.Count <= 0))
                return;

            DataGridViewSelectedRowCollection _DataGridViewSelectedRowCollection = ((DataGridView)sender).SelectedRows;
            DataGridViewRow _DataGridViewRow = (DataGridViewRow)_DataGridViewSelectedRowCollection[0];

            dgvFrom.ClearSelection();
            dgvFrom.Rows[_DataGridViewRow.Index].Selected = true;
        }

        private void ctnMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (rtxtLog.Text.Length != 0)
                switch (e.ClickedItem.Name.ToString())
                {
                    case "tsmCopy":
                        Clipboard.SetText(rtxtLog.SelectedText, TextDataFormat.Text);
                        break;

                    case "tsmSelectAll":
                        rtxtLog.SelectAll();
                        break;

                    case "tsmClear":
                        ClearLog();
                        break;

                    default:
                        break;
                }
        }

        private void CheckDirSearch()
        {
            lblSearchPattern.Visible = chkDir.Checked;
            ClearGrid();
            SetDefaultFilter();
            txtSourceRESXFile.Text = string.Empty;
            txtSearchPattern.Visible = chkDir.Checked;
            if (chkDir.Checked)
                lblSourceResxFile.Text = _DirSearch;
            else
                lblSourceResxFile.Text = _DirUnSearch;
        }

        private void InitLoading()
        {
            pnlMain.Enabled = false;
            CheckDirSearch();
            startupTimer = new Timer { Interval = 1000 };
            startupTimer.Tick += delegate
            {
                startupTimer.Enabled = false;
                Cursor = Cursors.WaitCursor;
                try
                {
                    pnlMain.Enabled = LoadLanguageToCombox();
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            };
            this.startupTimer.Start();

            TranslatorTextHelper _TranslatorTextHelper = new TranslatorTextHelper();
            txtTranslateAPIBaseUrl.Text = _TranslatorTextHelper.TranslateAPIBaseUrl;
            txtTranslateSubscriptionKey.Text = _TranslatorTextHelper.TranslateSubscriptionKey;
            chkSaveTranlsateSubscriptionKey.Checked = (!string.IsNullOrWhiteSpace(txtTranslateSubscriptionKey.Text));

            txtBingSpellCheckUrl.Text = _TranslatorTextHelper.BingSpellCheckUrl;
            txtBingSpellCheckSubscriptionKey.Text = _TranslatorTextHelper.BingSpellCheckSubscriptionKey;
            chkSaveBingSpellCheckSubscriptionKey.Checked = (!string.IsNullOrWhiteSpace(txtBingSpellCheckSubscriptionKey.Text));
            chkBingSpellCheck.Enabled = (!string.IsNullOrWhiteSpace(txtBingSpellCheckSubscriptionKey.Text));

            txtProxyAddress.Text = _TranslatorTextHelper.ProxyUrl;
            txtProxyUserName.Text = _TranslatorTextHelper.ProxyUserName;
            txtProxyPassword.Text = _TranslatorTextHelper.ProxyPassword;
            chkProxyUseDefaultCred.Checked = _TranslatorTextHelper.ProxyUseDefaultCredentials;
        }

        private bool LoadLanguageToCombox()
        {
            bool _LoadLanguageNoError = false;
            TranslatorTextHelper _TranslatorTextHelper = new TranslatorTextHelper();
            try
            {
                AddLogTextLine(Properties.Resources.LoadLanguageToCombox);
                try
                {
                    _TranslatorTextHelper.PopulateLanguageMenus(cboSourceFilter);
                    CopyComboLanguage(cboSourceFilter, cboFromLanguage, false, cboSourceFilter.SelectedIndex);
                    CopyComboLanguage(cboSourceFilter, cboToLanguage, true, 0);
                    _LoadLanguageNoError = true;
                }
                finally
                {
                    ClearStatusText();
                }
            }
            catch (Exception e)
            {
                TranslatorTextHelper.LogExceptionError(e, AddLogText);
            }
            return _LoadLanguageNoError;
        }

        private void SaveResourceFile()
        {
            ResourceHelper _ResourceHelper = new ResourceHelper();
            bool _SaveSuccessful = _ResourceHelper.SaveGridToResource(dgvTo, AddLogText);
            if (_SaveSuccessful)
                MessageBox.Show(Properties.Resources.FileSaveMessage, Properties.Resources.FileSaveMessageCaption,
                     MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SaveTranslateAPISettings()
        {
            TranslatorTextHelper _TranslatorTextHelper = new TranslatorTextHelper();
            _TranslatorTextHelper.TranslateAPIBaseUrl = txtTranslateAPIBaseUrl.Text;
            _TranslatorTextHelper.TranslateSubscriptionKey = txtTranslateSubscriptionKey.Text;
            _TranslatorTextHelper.SaveTranslateApiSettings(chkSaveTranlsateSubscriptionKey.Checked);
        }

        private void SaveBingSpellCheckSettings()
        {
            TranslatorTextHelper _TranslatorTextHelper = new TranslatorTextHelper();
            _TranslatorTextHelper.BingSpellCheckUrl = txtBingSpellCheckUrl.Text;
            _TranslatorTextHelper.BingSpellCheckSubscriptionKey = txtBingSpellCheckSubscriptionKey.Text;
            _TranslatorTextHelper.SaveBingSpellCheckSettings(chkSaveBingSpellCheckSubscriptionKey.Checked);
        }

        private void SaveProxySettings()
        {
            TranslatorTextHelper _TranslatorTextHelper = new TranslatorTextHelper();
            _TranslatorTextHelper.ProxyUrl = txtProxyAddress.Text;
            _TranslatorTextHelper.ProxyUserName = txtProxyUserName.Text;
            _TranslatorTextHelper.ProxyPassword = txtProxyPassword.Text;
            if (string.IsNullOrWhiteSpace(txtProxyAddress.Text))
                chkProxyUseDefaultCred.Checked = false;
            _TranslatorTextHelper.ProxyUseDefaultCredentials = chkProxyUseDefaultCred.Checked;
            _TranslatorTextHelper.SaveProxySettings();
        }

        private void TranslateLanguage()
        {
            if (dgvFrom.Rows.Count <= 0)
            {
                AddLogTextInternal(string.Format("\r\n{0}", Properties.Resources.NothingToTranslate));
                return;
            }
            if (btnTranslate.Text.Equals(Properties.Resources.ButtonTranslate))
            {
                if (MessageBox.Show(string.Format(Properties.Resources.ConfirmMessage, cboFromLanguage.Text, cboToLanguage.Text),
                    Properties.Resources.ConfirmMessageCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
                    return;
                btnTranslate.Text = Properties.Resources.ButtonStop;
                Translate();
            }
            else
            {
                btnTranslate.Text = Properties.Resources.ButtonTranslate;
                _AbortRequested = true;
            }
        }

        private void Translate()
        {
            dgvTo.Rows.Clear();
            Application.DoEvents();
            Cursor.Current = Cursors.WaitCursor;

            int _FileCount = 0;

            string _FileName = string.Empty;
            string _FromLanguage = GetLanguageTag(cboFromLanguage);
            string _FromCultureName = GetCultureName(cboFromLanguage);

            string _ToLanguage = GetLanguageTag(cboToLanguage);
            string _ToCultureName = GetCultureName(cboToLanguage);

            _AbortRequested = false;

            AddLogTextLine(string.Empty);
            AddLogTextLine(string.Format(Properties.Resources.TranslateingStart, cboFromLanguage.Text, _FromCultureName,
                cboToLanguage.Text, _ToCultureName));

            TranslationResult _TranslationResult = new TranslationResult();
            TranslatorTextHelper _TranslatorTextHelper = new TranslatorTextHelper(_FromLanguage, _ToLanguage, _FromCultureName);

            _TranslatorTextHelper.TranslateAPIBaseUrl = txtTranslateAPIBaseUrl.Text;
            _TranslatorTextHelper.TranslateSubscriptionKey = txtTranslateSubscriptionKey.Text;

            _TranslatorTextHelper.BingSpellCheckUrl = txtBingSpellCheckUrl.Text;
            _TranslatorTextHelper.BingSpellCheckSubscriptionKey = txtBingSpellCheckSubscriptionKey.Text;

            _TranslatorTextHelper.ProxyUrl = txtProxyAddress.Text;
            _TranslatorTextHelper.ProxyUserName = txtProxyUserName.Text;
            _TranslatorTextHelper.ProxyPassword = txtProxyPassword.Text;
            _TranslatorTextHelper.ProxyUseDefaultCredentials = chkProxyUseDefaultCred.Checked;

            foreach (DataGridViewRow _DataGridViewRow in dgvFrom.Rows)
            {
                if (!_FileName.Equals(_DataGridViewRow.Cells["FromFileName"].Value.ToString(), StringComparison.OrdinalIgnoreCase))
                    _FileCount += 1;

                _FileName = _DataGridViewRow.Cells["FromFileName"].Value.ToString();
                _TranslationResult += _TranslatorTextHelper.Translation(
                            _FromCultureName, _ToCultureName,
                            chkBingSpellCheck.Checked,
                            _FileCount,
                            _DataGridViewRow, dgvTo,
                            AddLogText, () => _AbortRequested);
                SetStatusText(string.Format(Properties.Resources.TranslatingLineNo, _DataGridViewRow.Index, _FileCount));
                if (_AbortRequested)
                    break;
            }
            btnTranslate.Text = Properties.Resources.ButtonTranslate;
            if (!_AbortRequested)
                AddLogTextLine(Properties.Resources.TranslateingEnd);
            else
                AddLogTextLine(Properties.Resources.TranslateingIncomplete);
            ClearStatusText();
        }

        private void LoadRESXFileToGrid()
        {
            ResourceHelper _ResourceHelper = new ResourceHelper();
            Cursor = Cursors.WaitCursor;

            ClearGrid();
            AddLogTextLine(string.Empty);
            try
            {
                if (chkDir.Checked)
                    _ResourceHelper.LoadMultiResourceToGrid(txtSourceRESXFile.Text, txtSearchPattern.Text, dgvFrom, AddLogTextLine);
                else
                    _ResourceHelper.LoadSingleResourceToGrid(txtSourceRESXFile.Text, dgvFrom, true, AddLogTextLine);
            }
            catch (Exception ex)
            {
                AddLogText(string.Format("Error :{0}", ex.Message));
            }
            finally
            {
                Cursor = Cursors.Default;
                ClearStatusText();
            }
        }

        private void AddLogText(string Messages)
        {
            if (Messages.Contains("Error"))
                _AbortRequested = true;

            AddLogTextInternal(Messages);
        }

        private void AddLogTextInternal(string Messages)
        {
            rtxtLog.Text += Messages;
            if (rtxtLog.TextLength != 0)
                rtxtLog.Text += "\r\n";
            rtxtLog.ScrollToCaret();
            rtxtLog.SelectionStart = rtxtLog.Text.Length;
            rtxtLog.Refresh();
        }

        private void AddLogTextLine(string Messages)
        {
            AddLogTextInternal(Messages);
            SetStatusText(Messages);
            _StatusLineText = Messages;
        }

        private void SetStatusText(string Messages)
        {
            toolStripLogStatusText.Text = Messages;
            statusStrip.Update();
        }

        private void ClearLog()
        {
            rtxtLog.Clear();
            ClearStatusText();
        }

        private void ClearStatusText()
        {
            _StatusLineText = string.Empty;
            toolStripLogStatusText.Text = string.Empty;
            statusStrip.Update();
        }

        private void ClearGrid()
        {
            dgvFrom.Rows.Clear();
            dgvTo.Rows.Clear();
        }

        private void SetDefaultFilter()
        {
            if (cboSourceFilter.Items.Count <= 0)
                txtSearchPattern.Text = string.Format("{0}.{1}", "*", _DefaultExtension);
            else
            {
                txtSearchPattern.Text = string.Format("*.{0}.{1}", GetCultureName(cboSourceFilter), _DefaultExtension);
                CopyComboLanguage(cboFromLanguage, cboToLanguage, true, 0);
            }
        }

        private void CopyComboLanguage(ComboBox CopyFromCombox, ComboBox CopyToCombox, bool RemoveSelected, int DefIndex)
        {
            CopyToCombox.Items.Clear();
            CopyToCombox.DisplayMember = "Key";
            CopyToCombox.ValueMember = "Value";

            if (CopyFromCombox.Items.Count <= 0)
                return;

            object[] _Objects = new object[CopyFromCombox.Items.Count];
            CopyFromCombox.Items.CopyTo(_Objects, 0);
            CopyToCombox.Items.AddRange(_Objects);
            if (RemoveSelected)
                CopyToCombox.Items.RemoveAt(CopyFromCombox.SelectedIndex);
            if (DefIndex >= 0)
                CopyToCombox.SelectedIndex = DefIndex;
        }

        private string GetLanguageTag(ComboBox Languages)
        {
            string LanguageTag = string.Empty;
            if (Languages.Items.Count <= 0)
                return LanguageTag;

            KeyValuePair<string, string> _KeyValuePair = (KeyValuePair<string, string>)Languages.SelectedItem;
            LanguageTag = _KeyValuePair.Value.ToString();
            return LanguageTag;
        }

        private string GetCultureName(ComboBox Languages)
        {
            string CultureName = string.Empty;
            if (Languages.Items.Count <= 0)
                return CultureName;

            KeyValuePair<string, string> _KeyValuePair = (KeyValuePair<string, string>)Languages.SelectedItem;
            CultureName = CultureHelper.GetCultureNameFromLetfLanguageTag(_KeyValuePair.Value.ToString());
            return CultureName;
        }
    }
}