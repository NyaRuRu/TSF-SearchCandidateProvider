using System.Globalization;
using TSF.InteropTypes;

namespace TSF.SearchCandidateProviderInternal
{
    internal static class LangIDEx
    {
        public static LangID ToLangID(this CultureInfo cultureInfo)
        {
            return new LangID((ushort)cultureInfo.LCID);
        }
        public static CultureInfo ToCultureInfo(this LangID langId)
        {
            return new CultureInfo((int)langId.LCID);
        }
    }
}
