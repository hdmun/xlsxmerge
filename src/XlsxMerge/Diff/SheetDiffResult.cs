namespace XlsxMerge.Diff;

public class SheetDiffResult
{
    public readonly string WorksheetName;
    public readonly ComparisonMode ComparisonMode;
    public readonly List<DocOrigin> DocsContaining; // 이 워크시트가 있는 문서.
    public List<DiffHunkInfo> HunkList = new();

    public SheetDiffResult(string worksheetName, ComparisonMode comparisonMode)
    {
        WorksheetName = worksheetName;
        ComparisonMode = comparisonMode;
        DocsContaining = new();
    }
}
