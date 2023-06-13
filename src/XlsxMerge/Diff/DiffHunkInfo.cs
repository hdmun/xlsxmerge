namespace XlsxMerge.Diff;

public class DiffHunkInfo
{
    public Diff3HunkStatus hunkStatus;
    public Dictionary<DocOrigin, RowRange> rowRangeMap = new Dictionary<DocOrigin, RowRange>();
}