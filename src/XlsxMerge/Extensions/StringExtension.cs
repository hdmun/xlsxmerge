using XlsxMerge.Features.Diffs;

namespace XlsxMerge.Extensions;

public static class StringExtension
{
    public static string AddDoubleQuote(this string s)
    {
        if (s.StartsWith("\"") && s.EndsWith("\""))
            return s;
        return "\"" + s + "\"";
    }

    public static DocOrigin? ToDocOrigin(this string s)
    {
        if (s.StartsWith("base"))
            return DocOrigin.Base;
        if (s.StartsWith("mine"))
            return DocOrigin.Mine;
        if (s.StartsWith("theirs"))
            return DocOrigin.Theirs;

        return null;
    }
}
