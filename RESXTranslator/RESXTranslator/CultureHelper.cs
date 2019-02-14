using System.Globalization;
using System.Linq;

namespace RESXTranslator
{
    internal class CultureHelper
    {
        private static CultureInfo[] _CultureInfos = CultureInfo.GetCultures(CultureTypes.NeutralCultures);

        private static string GetCultureName(CultureInfo CultureInfos)
        {
            if (CultureInfos == null)
                return string.Empty;

            return ((TextInfo)CultureInfos.TextInfo).CultureName;
        }

        public static string GetCultureNameFromLetfLanguageTag(string LetfLanguageTag)
        {
            CultureInfo _CultureInfo = CultureInfo.GetCultureInfoByIetfLanguageTag(LetfLanguageTag);
            return GetCultureName(_CultureInfo);
        }

        public static string GetCultureNameFromName(string Names)
        {
            CultureInfo _CultureInfo = CultureInfo.GetCultureInfo(Names);
            return GetCultureName(_CultureInfo);
        }

        public static string GetCultureNameFromEnglishName(string EnglishName)
        {
            CultureInfo _CultureInfo = _CultureInfos.FirstOrDefault(CultureInfo => CultureInfo.EnglishName.Equals(EnglishName));
            return GetCultureName(_CultureInfo);
        }
    }
}