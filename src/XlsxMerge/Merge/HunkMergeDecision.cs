using XlsxMerge.Diff;

namespace XlsxMerge.Merge;

public class HunkMergeDecision
{
    public List<DocOrigin> DocMergeOrder = new List<DocOrigin>(); // null = Conflict 상태로 둡니다. Empty = 모두 삭제. 
    public readonly DiffHunkInfo BaseHunkInfo;
    public List<List<DocOrigin>> DocMergeOrderCandidates = null; // 이 Hunk에서 선택 가능한 document merge orders.

    public HunkMergeDecision(DiffHunkInfo baseHunkInfo)
    {
        BaseHunkInfo = baseHunkInfo;
        BuildDocMergeOrderCandidates();
    }

    private void BuildDocMergeOrderCandidates()
    {
        DocMergeOrderCandidates = new List<List<DocOrigin>>();

        // add
        if (BaseHunkInfo.hunkStatus == Diff3HunkStatus.Conflict)
            DocMergeOrderCandidates.Add(null);

        DocMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Mine });
        DocMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Base });
        DocMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Theirs });
        DocMergeOrderCandidates.Add(new List<DocOrigin>()); // "Delete" Menu

        DocMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Mine, DocOrigin.Base });
        DocMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Mine, DocOrigin.Base, DocOrigin.Theirs });
        DocMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Mine, DocOrigin.Theirs });
        DocMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Mine, DocOrigin.Theirs, DocOrigin.Base });

        DocMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Base, DocOrigin.Mine });
        DocMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Base, DocOrigin.Mine, DocOrigin.Theirs });
        DocMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Base, DocOrigin.Theirs });
        DocMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Base, DocOrigin.Theirs, DocOrigin.Mine });

        DocMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Theirs, DocOrigin.Base });
        DocMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Theirs, DocOrigin.Base, DocOrigin.Mine });
        DocMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Theirs, DocOrigin.Mine });
        DocMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Theirs, DocOrigin.Mine, DocOrigin.Base });

        // exclude
        List<DocOrigin> docOriginsToExclude = new List<DocOrigin>();
        switch (BaseHunkInfo.hunkStatus)
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
        var rowRangeMap = BaseHunkInfo.rowRangeMap;
        foreach (var docOrigin in new[] { DocOrigin.Base, DocOrigin.Mine, DocOrigin.Theirs })
            if (rowRangeMap.ContainsKey(docOrigin) == false || rowRangeMap[docOrigin].RowCount == 0)
                docOriginsToExclude.Add(docOrigin);

        foreach (var docOrigin in docOriginsToExclude)
            DocMergeOrderCandidates.RemoveAll(r => r != null && r.Contains(docOrigin));
    }
}
