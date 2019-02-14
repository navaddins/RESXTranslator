using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Windows.Forms;

namespace RESXTranslator
{
    public class TranslatorTextHelper
    {
        private const string _DefaultTranslateAPIUrl = "https://api.cognitive.microsofttranslator.com/{0}?api-version=3.0";
        private const string _DefaultBingSpellCheckUrl = "https://api.cognitive.microsoft.com/bing/v7.0/spellcheck/";
        private const string _ResxExtension = ".resx";

        private string _TranslateAPIBaseUrl = string.Empty;
        private string _TranslateSubscriptionKey = string.Empty;

        private string _BingSpellCheckUrl = string.Empty;
        private string _BingSpellCheckSubscriptionKey = string.Empty;

        private static string _ProxyUrl = string.Empty;
        private static string _ProxyUserName = string.Empty;
        private static string _ProxyPassword = string.Empty;
        private static bool _ProxyUseDefaultCredentials = false;

        private HttpClient _HttpClient = null;

        private HttpClient HttpClients
        {
            get
            {
                return _HttpClient = _HttpClient ?? new HttpClient(HttpClientHandlers);
            }
        }

        private HttpClientHandler HttpClientHandlers
        {
            get
            {
                return GetHttpClientHandler();
            }
        }

        private string[] _LanguageCodes;

        private string TranslateFrom { get; set; }

        private string TranslateTo { get; set; }

        private string SpellCheckCulture { get; set; }

        public string TranslateAPIBaseUrl
        {
            get
            {
                return _TranslateAPIBaseUrl;
            }
            set
            {
                _TranslateAPIBaseUrl = value;
            }
        }

        public string TranslateSubscriptionKey
        {
            get
            {
                return _TranslateSubscriptionKey;
            }
            set
            {
                _TranslateSubscriptionKey = value;
            }
        }

        public string BingSpellCheckUrl
        {
            get
            {
                return _BingSpellCheckUrl;
            }
            set
            {
                _BingSpellCheckUrl = value;
            }
        }

        public string BingSpellCheckSubscriptionKey
        {
            get
            {
                return _BingSpellCheckSubscriptionKey;
            }
            set
            {
                _BingSpellCheckSubscriptionKey = value;
            }
        }

        public string ProxyUrl
        {
            get
            {
                return _ProxyUrl;
            }
            set
            {
                _ProxyUrl = value;
            }
        }

        public string ProxyUserName
        {
            get
            {
                return _ProxyUserName;
            }
            set
            {
                _ProxyUserName = value;
            }
        }

        public string ProxyPassword
        {
            get
            {
                return _ProxyPassword;
            }
            set
            {
                _ProxyPassword = value;
            }
        }

        public bool ProxyUseDefaultCredentials
        {
            get
            {
                return _ProxyUseDefaultCredentials;
            }
            set
            {
                _ProxyUseDefaultCredentials = value;
            }
        }

        public delegate void LogEvent(string Messages);

        public delegate bool ShallAbort();

        public Uri GetLanguageUrl
        {
            get
            {
                return new Uri(string.Format("{0}&{1}", string.Format(TranslateAPIBaseUrl, "languages"), "scope=translation"));
            }
        }

        public Uri GetBingSpellCheckUrl
        {
            get
            {
                return new Uri(String.Format("{0}{1}", _DefaultBingSpellCheckUrl, string.Format("?mode=spell&mkt={0}", SpellCheckCulture)));
            }
        }

        public Uri GetTranslateUrl
        {
            get
            {
                return new Uri(string.Format("{0}&{1}", string.Format(TranslateAPIBaseUrl, "translate"), string.Format("scope=translation&from={0}&to={1}", TranslateFrom, TranslateTo)));
            }
        }

        public TranslatorTextHelper()
        {
            GetTranslateApiSettings();
            GetBingSpellCheckSettings();
            GetProxySettings();
        }

        public TranslatorTextHelper(string FromLanguage, string ToLanguage, string SpellCheckCultureName)
        {
            GetTranslateApiSettings();
            GetBingSpellCheckSettings();
            GetProxySettings();
            TranslateFrom = FromLanguage;
            TranslateTo = ToLanguage;
            SpellCheckCulture = SpellCheckCultureName;
        }

        public void PopulateLanguageMenus(ComboBox Languages)
        {
            SortedDictionary<string, string> _LanguageCodesAndTitles = GetAvailableLanguages();
            int count = _LanguageCodesAndTitles.Count;
            foreach (KeyValuePair<string, string> _LanguageCodesAndTitle in _LanguageCodesAndTitles)
                Languages.Items.Add(_LanguageCodesAndTitle);

            Languages.DisplayMember = "Key";
            Languages.ValueMember = "Value";
            Languages.Text = "English";
        }

        public TranslationResult Translation(
            string FromLanguageCultureName, string ToLanguageCultureName,
            bool SpellCheck,
            int FileCount,
            DataGridViewRow SourceResxGridViewRow,
            DataGridView ResxGridView,
            LogEvent LogEvents,
            ShallAbort AbortRequested)
        {
            TranslationResult _TranslationResult = new TranslationResult();

            if ((AbortRequested != null) && AbortRequested())
                return _TranslationResult;

            string _TextToTranslate = SourceResxGridViewRow.Cells["FromValue"].Value.ToString();
            string _TranslatedText = string.Empty;

            try
            {
                try
                {
                    if (_TextToTranslate.Length > int.Parse(Properties.Resources.StringMaxLength))
                        throw new Exception(Properties.Resources.StringMaxLengthError);

                    if (SpellCheck)
                        _TranslatedText = CorrectSpelling(_TextToTranslate);
                    _TranslatedText = TranslateLanguage(_TextToTranslate);
                }
                catch (Exception e)
                {
                    LogExceptionError(e, LogEvents);
                    return _TranslationResult;
                }

                Application.DoEvents();
            }
            finally
            {
            }

            _TranslationResult.FileCount = FileCount;
            _TranslationResult.StringCount = _TextToTranslate.Length;

            _TranslationResult.TranslatFromString = _TextToTranslate;
            _TranslationResult.TranslateFromLanguage = TranslateFrom;
            _TranslationResult.TranslateFromCultureName = FromLanguageCultureName;

            _TranslationResult.TranslatedString = _TranslatedText;
            _TranslationResult.TranslatedToLanguage = TranslateTo;
            _TranslationResult.TranslatedCultureName = ToLanguageCultureName;

            AddTranslatedTextToGrid(SourceResxGridViewRow, ResxGridView, _TranslationResult, LogEvents);
            return _TranslationResult;
        }

        public void SaveTranslateApiSettings(bool SaveSubscriptionKey)
        {
            Properties.Settings.Default.TranslateAPIBaseUrl = TranslateAPIBaseUrl;
            Properties.Settings.Default.TranslateSubscriptionKey = string.Empty;
            if (SaveSubscriptionKey)
                Properties.Settings.Default.TranslateSubscriptionKey = TranslateSubscriptionKey;
            Properties.Settings.Default.Save();
        }

        public void SaveBingSpellCheckSettings(bool SaveSubscriptionKey)
        {
            Properties.Settings.Default.BingSpellCheckUrl = BingSpellCheckUrl;
            Properties.Settings.Default.BingSpellCheckSubscriptionKey = string.Empty;
            if (SaveSubscriptionKey)
                Properties.Settings.Default.BingSpellCheckSubscriptionKey = BingSpellCheckSubscriptionKey;
            Properties.Settings.Default.Save();
        }

        public void SaveProxySettings()
        {
            Properties.Settings.Default.ProxyUrl = ProxyUrl;
            Properties.Settings.Default.ProxyUserName = ProxyUserName;
            Properties.Settings.Default.ProxyPassword = ProxyPassword;
            Properties.Settings.Default.ProxyUseDefaultCredentials = ProxyUseDefaultCredentials;
            Properties.Settings.Default.Save();
        }

        public static void LogExceptionError(Exception ex, LogEvent LogEvents)
        {
            if (ex is ArgumentOutOfRangeException)
            {
                if (LogEvents != null)
                {
                    if (ex.InnerException != null)
                    {
                        if (ex.InnerException.InnerException != null)
                            LogExceptionError(ex.InnerException.InnerException, LogEvents);
                        else
                            LogExceptionError(ex.InnerException, LogEvents);
                    }
                    else
                        LogEvents("Error: " + ex.Message);
                }
            }
            else if (ex is WebException)
            {
                if (LogEvents != null)
                {
                    if (ex.InnerException != null)
                    {
                        if (ex.InnerException.InnerException != null)
                            LogExceptionError(ex.InnerException.InnerException, LogEvents);
                        else
                            LogExceptionError(ex.InnerException, LogEvents);
                    }
                    else
                        LogEvents("Microsoft Translator API Error: " + ex.Message);
                }
            }
            else
            {
                if (LogEvents != null)
                {
                    if (ex.InnerException != null)
                    {
                        if (ex.InnerException.InnerException != null)
                            LogExceptionError(ex.InnerException.InnerException, LogEvents);
                        else
                            LogExceptionError(ex.InnerException, LogEvents);
                    }
                    else
                        LogEvents("General Error: " + ex.Message);
                }
            }
        }

        private string CorrectSpelling(string TextToSpellCheck)
        {
            if (string.IsNullOrWhiteSpace(TextToSpellCheck))
                return TextToSpellCheck;

            using (HttpRequestMessage _HttpRequestMessage = GetHttpRequestMessage(HttpMethod.Post, GetBingSpellCheckUrl, BingSpellCheckSubscriptionKey))
            {
                Dictionary<string, string> _TextParameter = new Dictionary<string, string> { { Properties.Resources.MSBingText, TextToSpellCheck } };
                _HttpRequestMessage.Content = new FormUrlEncodedContent(_TextParameter);
                _HttpRequestMessage.Content.Headers.ContentType = new MediaTypeHeaderValue(Properties.Resources.RequestHeaderContentUrlEncoded);
                _HttpRequestMessage.Content.Headers.ContentType.CharSet = Properties.Resources.RequestHeaderAcceptCharSet;

                HttpResponseMessage _HttpResponseMessage = HttpClients.SendAsync(_HttpRequestMessage).Result;
                _HttpResponseMessage.EnsureSuccessStatusCode();

                JObject _JsonResponse = JObject.Parse(JsonConvert.DeserializeObject(_HttpResponseMessage.Content.ReadAsStringAsync().Result).ToString());
                JArray _FlaggedTokens = (JArray)_JsonResponse[Properties.Resources.MSBingFlaggedTokens];

                SortedDictionary<int, string[]> _SortedDictionary = new SortedDictionary<int, string[]>(Comparer.Create<int>((a, b) => b.CompareTo(a)));
                for (int i = 0; i < _FlaggedTokens.Count; i++)
                {
                    JToken _Correction = _FlaggedTokens[i];
                    JToken _Suggestion = _Correction[Properties.Resources.MSBingSuggestions][0];
                    if (decimal.Parse((_Suggestion[Properties.Resources.MSBingScore].ToString())) > (decimal)0.7)
                        _SortedDictionary[(int)_Correction[Properties.Resources.MSBingOffset]] =
                            new string[] { _Correction[Properties.Resources.MSBingToken].ToString(), _Suggestion[Properties.Resources.MSBingSuggestion].ToString() };  // dict value = {error, correction}
                }

                // apply the corrections in order from right to left
                foreach (int i in _SortedDictionary.Keys)
                {
                    string _OldSpelling = _SortedDictionary[i][0];
                    string _NewSpelling = _SortedDictionary[i][1];

                    // apply capitalization from original text to correction - all caps or initial caps
                    if (TextToSpellCheck.Substring(i, _OldSpelling.Length).All(char.IsUpper))
                        _NewSpelling = _NewSpelling.ToUpper();
                    else if (char.IsUpper(TextToSpellCheck[i]))
                        _NewSpelling = _NewSpelling[0].ToString().ToUpper() + _NewSpelling.Substring(1);

                    TextToSpellCheck = TextToSpellCheck.Substring(0, i) + _NewSpelling + TextToSpellCheck.Substring(i + _OldSpelling.Length);
                }
            }
            return TextToSpellCheck;
        }

        private string TranslateLanguage(string TextToTranslate)
        {
            if (string.IsNullOrWhiteSpace(TextToTranslate))
                return TextToTranslate;

            object[] _objBody = new object[] { new { Text = TextToTranslate } };
            string _RequestBody = JsonConvert.SerializeObject(_objBody);

            using (HttpRequestMessage _HttpRequestMessage = GetHttpRequestMessage(HttpMethod.Post, GetTranslateUrl, TranslateSubscriptionKey))
            {
                _HttpRequestMessage.Content = new StringContent(_RequestBody, Encoding.UTF8, Properties.Resources.RequestHeaderContentJson);
                HttpResponseMessage _HttpResponseMessage = HttpClients.SendAsync(_HttpRequestMessage).Result;
                _HttpResponseMessage.EnsureSuccessStatusCode();

                List<Dictionary<string, List<Dictionary<string, string>>>> _ListDictionary =
                    JsonConvert.DeserializeObject<List<Dictionary<string, List<Dictionary<string, string>>>>>(_HttpResponseMessage.Content.ReadAsStringAsync().Result);
                List<Dictionary<String, String>> _Translations = _ListDictionary[0][Properties.Resources.MSAPITranslations];
                if (!TextToTranslate.ToLower().Equals(_Translations[0][Properties.Resources.MSAPITranslationText].ToString().ToLower()))
                    TextToTranslate = _Translations[0][Properties.Resources.MSAPITranslationText].ToString();
            }
            return TextToTranslate;
        }

        private void GetTranslateApiSettings()
        {
            TranslateAPIBaseUrl = Properties.Settings.Default.TranslateAPIBaseUrl;
            TranslateSubscriptionKey = Properties.Settings.Default.TranslateSubscriptionKey;
            if (string.IsNullOrWhiteSpace(TranslateAPIBaseUrl))
                TranslateAPIBaseUrl = _DefaultTranslateAPIUrl;
        }

        private void GetBingSpellCheckSettings()
        {
            BingSpellCheckUrl = Properties.Settings.Default.BingSpellCheckUrl;
            BingSpellCheckSubscriptionKey = Properties.Settings.Default.BingSpellCheckSubscriptionKey;
            if (string.IsNullOrWhiteSpace(BingSpellCheckUrl))
                BingSpellCheckUrl = _DefaultBingSpellCheckUrl;
        }

        private void GetProxySettings()
        {
            ProxyUrl = Properties.Settings.Default.ProxyUrl;
            ProxyUserName = Properties.Settings.Default.ProxyUserName;
            ProxyPassword = Properties.Settings.Default.ProxyPassword;
            ProxyUseDefaultCredentials = Properties.Settings.Default.ProxyUseDefaultCredentials;
        }

        private SortedDictionary<string, string> GetAvailableLanguages()
        {
            SortedDictionary<string, string> _LanguageCodesAndTitles =
                new SortedDictionary<string, string>(Comparer.Create<string>((a, b) => string.Compare(a, b, true)));

            using (HttpRequestMessage _HttpRequestMessage = GetHttpRequestMessage(HttpMethod.Get, GetLanguageUrl, TranslateSubscriptionKey))
            {
                HttpResponseMessage _HttpResponseMessage = HttpClients.SendAsync(_HttpRequestMessage).Result;
                _HttpResponseMessage.EnsureSuccessStatusCode();

                JObject _JObject = JObject.Parse(_HttpResponseMessage.Content.ReadAsStringAsync().Result);
                Dictionary<string, Dictionary<string, string>> _Languages =
                    JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>>>(_JObject[Properties.Resources.MSAPITranslation].ToString());
                _LanguageCodes = _Languages.Keys.ToArray();
                foreach (KeyValuePair<string, Dictionary<string, string>> _KeyValue in _Languages)
                    _LanguageCodesAndTitles.Add(_KeyValue.Value[Properties.Resources.MSAPITranslationLangName], _KeyValue.Key);
            }
            return _LanguageCodesAndTitles;
        }

        private void AddTranslatedTextToGrid(DataGridViewRow SourceResxGridViewRow, DataGridView ResxGridView, TranslationResult TranslatedResult,
            LogEvent LogEvents)
        {
            string _FromResxFilePath = SourceResxGridViewRow.Cells["FromFileName"].Value.ToString();
            string _Directory = Path.GetDirectoryName(_FromResxFilePath);
            string _ResxFromFileExtension = Path.GetExtension(_FromResxFilePath);
            string _ResxRootFileName = Path.GetFileNameWithoutExtension(_FromResxFilePath);
            string _OrigLangTag = Path.GetExtension(_ResxRootFileName);
            string _ToResxPath;

            if (!string.IsNullOrWhiteSpace(TranslatedResult.TranslateFromCultureName))
                _ToResxPath = Path.ChangeExtension(_ResxRootFileName, TranslatedResult.TranslatedCultureName + _ResxFromFileExtension);
            else
                _ToResxPath = _ResxRootFileName + "." + TranslatedResult.TranslatedCultureName + _ResxFromFileExtension;

            if (!string.IsNullOrEmpty(_Directory))
                _ToResxPath = Path.Combine(_Directory, _ToResxPath);

            DataGridViewRow _DataGridViewRow = new DataGridViewRow();
            _DataGridViewRow.CreateCells(ResxGridView);
            _DataGridViewRow.SetValues(new object[] { SourceResxGridViewRow.Cells["FromKey"].Value.ToString(), TranslatedResult.TranslatedString, _ToResxPath });
            ResxGridView.Rows.Add(_DataGridViewRow);
            ResxGridView.Refresh();
        }

        private static HttpRequestMessage GetHttpRequestMessage(HttpMethod HttpMethods, Uri RequestUri, string SubscriptionKey)
        {
            HttpRequestMessage _HttpRequestMessage = new HttpRequestMessage();
            _HttpRequestMessage.RequestUri = RequestUri;
            _HttpRequestMessage.Method = HttpMethods;
            _HttpRequestMessage.Headers.Clear();
            _HttpRequestMessage.Headers.AcceptLanguage.Add(new StringWithQualityHeaderValue(Properties.Resources.RequestHeaderAcceptedLanguage));
            _HttpRequestMessage.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue(Properties.Resources.RequestHeaderContentJson));
            _HttpRequestMessage.Headers.AcceptCharset.Add(new StringWithQualityHeaderValue(Properties.Resources.RequestHeaderAcceptCharSet));
            _HttpRequestMessage.Headers.Add(Properties.Resources.RequestHeaderSubscriptionKey, SubscriptionKey);
            return _HttpRequestMessage;
        }

        private static HttpClientHandler GetHttpClientHandler()
        {
            IWebProxy _IWebProxy = GetWebProxy();
            HttpClientHandler _HttpClientHandle = new HttpClientHandler();
            _HttpClientHandle.Proxy = _IWebProxy;
            _HttpClientHandle.UseProxy = (_IWebProxy != null);
            return _HttpClientHandle;
        }

        private static WebProxy GetWebProxy()
        {
            WebProxy _WebProxy = null;
            if (string.IsNullOrWhiteSpace(_ProxyUrl.ToString()))
                return _WebProxy;

            _WebProxy = new WebProxy(new Uri(_ProxyUrl), false);
            _WebProxy.UseDefaultCredentials = (_ProxyUseDefaultCredentials);
            if (!_WebProxy.UseDefaultCredentials)
                _WebProxy.Credentials = new System.Net.NetworkCredential(_ProxyUserName, _ProxyPassword);
            return _WebProxy;
        }
    }
}