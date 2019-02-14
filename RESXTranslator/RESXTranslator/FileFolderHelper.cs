using System.Windows.Forms;

namespace RESXTranslator
{
    internal class FileFolderHelper
    {
        private string _DefaultFileName = string.Empty;
        private string _DefaultExtension = string.Empty;
        private string _DefaultFilter = Properties.Resources.DefaultAllFilter;
        private string _FolderBrowserDesc = string.Empty;

        public FileFolderHelper()
        {
        }

        public string OpenFileDialog(string PreviousFileName, string DialogTitle)
        {
            string _FileName = string.Empty;
            OpenFileDialog _OpenFileDialog = new OpenFileDialog
            {
                Title = DialogTitle,

                CheckFileExists = true,
                CheckPathExists = true,
                FileName = _DefaultFileName,
                DefaultExt = _DefaultExtension,
                Filter = _DefaultFilter,
                FilterIndex = 2,
                RestoreDirectory = true,
                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (_OpenFileDialog.ShowDialog() == DialogResult.OK)
                _FileName = _OpenFileDialog.FileName;
            return ((string.IsNullOrWhiteSpace(_FileName)) ? PreviousFileName : _FileName); ;
        }

        public string OpenFileDialog(string PreviousFileName, string DialogTitle, string DefaultFileName, string DefaultExtension, string DefaultFilter)
        {
            _DefaultFileName = DefaultFileName;
            _DefaultExtension = DefaultExtension;
            _DefaultFilter = DefaultFilter;
            return OpenFileDialog(PreviousFileName, DialogTitle);
        }

        public string FolderBrowserDialog(string PreviousFolderName)
        {
            string _FolderName = string.Empty;
            FolderBrowserDialog _FolderBrowserDialog = new FolderBrowserDialog
            {
                ShowNewFolderButton = false,
                Description = _FolderBrowserDesc
            };

            if (_FolderBrowserDialog.ShowDialog() == DialogResult.OK)
                _FolderName = _FolderBrowserDialog.SelectedPath;
            return ((string.IsNullOrWhiteSpace(_FolderName)) ? PreviousFolderName : _FolderName);
        }

        public string FolderBrowserDialog(string PreviousFolderName, string FolderBrowserDesc)
        {
            _FolderBrowserDesc = FolderBrowserDesc;
            return FolderBrowserDialog(PreviousFolderName);
        }
    }
}