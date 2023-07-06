using XlsxMerge.Features.Diffs;

namespace XlsxMerge.Extensions;

public static class ComparisonModeExtension
{
    public static IEnumerable<string> GetWorksheetColumns(this ComparisonMode comparisonMode)
    {
        yield return "워크시트 이름";
        if (comparisonMode == ComparisonMode.ThreeWay)
            yield return "충돌";

        yield return "Mine (Dest/Curr)";
        if (comparisonMode == ComparisonMode.ThreeWay)
            yield return "Theirs (Src/Others)";
    }
}
