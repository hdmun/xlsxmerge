namespace XlsxMerge.Diff;

public class SheetDiffResult
{
    public ComparisonMode ComparisonMode = ComparisonMode.Unknown;
    public string WorksheetName = "";
    public List<DocOrigin> DocsContaining = new List<DocOrigin>(); // 이 워크시트가 있는 문서.
    public List<DiffHunkInfo> HunkList = new List<DiffHunkInfo>();
}
