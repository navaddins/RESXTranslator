using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Resources;
using System.Windows.Forms;

namespace RESXTranslator
{
    internal class ResourceHelper
    {
        private List<DictionaryEntry> ReadResource(string ResxFilePath)
        {
            if (string.IsNullOrWhiteSpace(ResxFilePath))
                return new List<DictionaryEntry>();

            ResXResourceReader _ResXResourceReader = new ResXResourceReader(ResxFilePath);
            List<DictionaryEntry> _ResxDictionaryEntry = _ResXResourceReader.Cast<DictionaryEntry>().ToList();
            return _ResxDictionaryEntry;
        }

        public delegate void ProcessStatus(string Messages);

        public void LoadMultiResourceToGrid(string DirectoryRoot, string SearchPattern, DataGridView ResxGridView, ProcessStatus SetStatusText)
        {
            ResxGridView.Rows.Clear();

            if (!Directory.Exists(DirectoryRoot))
                throw new Exception(string.Format(Properties.Resources.DirectoryNotFound, DirectoryRoot));

            DirectoryInfo _DirectoryInfo = new DirectoryInfo(DirectoryRoot);
            FileInfo[] _FileInfos = _DirectoryInfo.GetFiles(SearchPattern, SearchOption.AllDirectories);
            if (_FileInfos.Length != 0)
                foreach (FileInfo _FileInfo in _FileInfos)
                    LoadSingleResourceToGrid(_FileInfo.FullName, ResxGridView, false, SetStatusText);
            else
                throw new Exception(Properties.Resources.FileNotFoundWithFilter);

            if (SetStatusText != null)
                SetStatusText(string.Format(Properties.Resources.TotalFileCountRead, _FileInfos.Length, DirectoryRoot));
        }

        public void LoadSingleResourceToGrid(string ResxFilePath, DataGridView ResxGridView, bool IsClearRows, ProcessStatus SetStatusText)
        {
            if (IsClearRows)
                ResxGridView.Rows.Clear();

            if (!File.Exists(ResxFilePath))
                throw new Exception(string.Format(Properties.Resources.FileNotFound, ResxFilePath));

            List<DictionaryEntry> _ResxDictionaryEntry = ReadResource(ResxFilePath);
            foreach (DictionaryEntry _DictionaryEntry in _ResxDictionaryEntry)
            {
                DataGridViewRow _DataGridViewRow = new DataGridViewRow();                
                _DataGridViewRow.CreateCells(ResxGridView);
                _DataGridViewRow.SetValues(new object[] { _DictionaryEntry.Key.ToString(), _DictionaryEntry.Value.ToString(), ResxFilePath });
                _DataGridViewRow.HeaderCell.Value = String.Format("{0}.", ResxGridView.Rows.Count + 1);
                ResxGridView.Rows.Add(_DataGridViewRow);                
            }
            ResxGridView.Refresh();
            Application.DoEvents();
            if (SetStatusText != null)
                SetStatusText(string.Format(Properties.Resources.ReadingSourceResxFile, ResxFilePath));
        }

        public bool SaveGridToResource(DataGridView ResxGridView, ProcessStatus SetStatusText)
        {
            SetStatusText(string.Empty);
            bool _SaveSuccessful = true;
            bool _IsOpening = false;
            string _FileName = string.Empty;
            Stream _Stream = null;
            ResXResourceWriter _ResxResourceWriter = null;

            try
            {
                if (ResxGridView.Rows.Count <= 0)
                    throw new Exception(Properties.Resources.NothingToSave);

                foreach (DataGridViewRow _DataGridViewRow in ResxGridView.Rows)
                {
                    if (!_FileName.Equals(_DataGridViewRow.Cells["ToFileName"].Value.ToString()))
                    {
                        _FileName = _DataGridViewRow.Cells["ToFileName"].Value.ToString();
                        if (File.Exists(_FileName))
                        {
                            try
                            {
                                File.Delete(_FileName);
                            }
                            catch (IOException)
                            {
                                _SaveSuccessful = false;
                                throw;
                            }
                            catch (UnauthorizedAccessException)
                            {
                                _SaveSuccessful = false;
                                throw;
                            }
                        }
                        if (!File.Exists(_FileName))
                        {
                            try
                            {
                                if (_IsOpening)
                                {
                                    _ResxResourceWriter.Close();
                                    _Stream.Close();
                                }
                                _Stream = new FileStream(_DataGridViewRow.Cells["ToFileName"].Value.ToString(), FileMode.CreateNew);
                                _ResxResourceWriter = new ResXResourceWriter(_Stream);
                                _IsOpening = true;
                                if (SetStatusText != null)
                                    SetStatusText(string.Format("{0}\r\n", string.Format(Properties.Resources.SaveResourceFileMessage, _FileName)));
                            }
                            catch (IOException)
                            {
                                _SaveSuccessful = false;
                                throw;
                            }
                            catch (UnauthorizedAccessException)
                            {
                                _SaveSuccessful = false;
                                throw;
                            }
                        }
                    }
                    _ResxResourceWriter.AddResource(_DataGridViewRow.Cells["ToKey"].Value.ToString(), _DataGridViewRow.Cells["ToValue"].Value.ToString());
                }
            }
            catch (IOException e)
            {
                if (SetStatusText != null)
                    SetStatusText(string.Format("{0}\r\n", e.Message));
                _SaveSuccessful = false;
            }
            catch (UnauthorizedAccessException e)
            {
                if (SetStatusText != null)
                    SetStatusText(string.Format("{0}\r\n", e.Message));
                _SaveSuccessful = false;
            }
            catch (Exception e)
            {
                if (SetStatusText != null)
                    SetStatusText(string.Format("{0}\r\n", e.Message));
                _SaveSuccessful = false;
            }
            finally
            {
                if (_IsOpening)
                {
                    _ResxResourceWriter.Close();
                    _ResxResourceWriter.Dispose();
                    _Stream.Close();
                    _Stream.Dispose();
                }
            }
            return _SaveSuccessful;
        }
    }
}