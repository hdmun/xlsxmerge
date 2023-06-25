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