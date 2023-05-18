namespace XlsxMerge.Extensions;

public static class StringExtension
{
    public static string AddDoubleQuote(this string s)
    {
        if (s.StartsWith("\"") && s.EndsWith("\""))
            return s;
        return "\"" + s + "\"";
    }
}
