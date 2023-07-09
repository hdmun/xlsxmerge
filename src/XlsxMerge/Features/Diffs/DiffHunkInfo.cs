namespace XlsxMerge.Features.Diffs;

public class DiffHunkInfo
{
    public readonly Diff3HunkStatus hunkStatus;
    public readonly Dictionary<DocOrigin, RowRange> rowRangeMap;

    public DiffHunkInfo(Diff3HunkStatus status)
    {
        hunkStatus = status;
        rowRangeMap = new();
    }

    public void Add(DocOrigin docOrigin, RowRange rowRange)
    {
        rowRangeMap.Add(docOrigin, rowRange);
    }

    public RowRange? GetRowRange(DocOrigin docOrigin)
    {
        return rowRangeMap.TryGetValue(docOrigin, out var rowRange) ? rowRange : null;
    }

    public int GetRowNumber(DocOrigin docOrigin)
    {
        var rowRange = GetRowRange(docOrigin);
        return rowRange?.RowNumber ?? 0;
    }

    public int GetRowCount(DocOrigin docOrigin)
    {
        var rowRange = GetRowRange(docOrigin);
        return rowRange?.RowCount ?? 0;
    }

    public IEnumerable<DocOrigin> ToExcludeDocOrigins()
    {
        var docOriginsToExclude = new HashSet<DocOrigin>();
        switch (hunkStatus)
        {
            case Diff3HunkStatus.BaseDiffers:
                docOriginsToExclude.Add(DocOrigin.Theirs);
                break;
            case Diff3HunkStatus.MineDiffers:
                docOriginsToExclude.Add(DocOrigin.Theirs);
                break;
            case Diff3HunkStatus.TheirsDiffers:
                docOriginsToExclude.Add(DocOrigin.Mine);
                break;
            case Diff3HunkStatus.Conflict:
                break;
        }

        var docOrigins = Enum.GetValues<DocOrigin>()
            .Where(x => rowRangeMap.ContainsKey(x) || rowRangeMap[x].RowCount == 0);
        foreach (var docOrigin in docOrigins)
        {
            docOriginsToExclude.Add(docOrigin);
        }

        return docOriginsToExclude;
    }
}