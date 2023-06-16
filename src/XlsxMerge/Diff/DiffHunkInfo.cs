namespace XlsxMerge.Diff;

public class DiffHunkInfo
{
    public readonly Diff3HunkStatus hunkStatus;
    public readonly Dictionary<DocOrigin, RowRange> rowRangeMap;

    public DiffHunkInfo(Diff3HunkStatus status)
    {
        hunkStatus = status;
        rowRangeMap = new();
    }
}