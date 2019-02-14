namespace RESXTranslator
{
    public struct TranslationResult
    {
        public int FileCount { get; set; }

        public int StringCount { get; set; }

        public string TranslatFromString { get; set; }

        public string TranslateFromLanguage { get; set; }

        public string TranslateFromCultureName { get; set; }

        public string TranslatedString { get; set; }

        public string TranslatedToLanguage { get; set; }

        public string TranslatedCultureName { get; set; }

        public static TranslationResult operator +(TranslationResult r1, TranslationResult r2)
        {
            return new TranslationResult
            {
                FileCount = r1.FileCount + r2.FileCount,
                StringCount = r1.StringCount + r2.StringCount,
            };
        }
    }
}