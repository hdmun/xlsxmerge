using XlsxMerge.Features.Diffs;
using XlsxMerge.Features.Diffs.Enums;

namespace XlsxMerge.Features.Merges;

public class HunkMergeDecision
{
    public List<DocOrigin>? DocMergeOrder; // null = Conflict 상태로 둡니다. Empty = 모두 삭제. 
    public readonly DiffHunkInfo BaseHunkInfo;
    public readonly List<List<DocOrigin>> DocMergeOrderCandidates; // 이 Hunk에서 선택 가능한 document merge orders.

    public HunkMergeDecision(DiffHunkInfo baseHunkInfo)
    {
        // bse에 비해 mine/theirs 모두 변경사항이 있고, 그 둘이 같으면 Mine을 사용한다.
        var docMergeOrder = new List<DocOrigin>();
        switch (baseHunkInfo.hunkStatus)
        {
            case Diff3HunkStatus.BaseDiffers:
            case Diff3HunkStatus.MineDiffers:
                docMergeOrder.Add(DocOrigin.Mine);
                break;
            case Diff3HunkStatus.TheirsDiffers:
                docMergeOrder.Add(DocOrigin.Theirs);
                break;
            case Diff3HunkStatus.Conflict:
                docMergeOrder = null;
                break;
        }

        DocMergeOrder = docMergeOrder;

        BaseHunkInfo = baseHunkInfo;

        var docMergeOrderCandidates = new List<List<DocOrigin>>();

        // add
        if (baseHunkInfo.hunkStatus == Diff3HunkStatus.Conflict)
            docMergeOrderCandidates.Add(null);

        docMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Mine });
        docMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Base });
        docMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Theirs });
        docMergeOrderCandidates.Add(new List<DocOrigin>()); // "Delete" Menu

        docMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Mine, DocOrigin.Base });
        docMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Mine, DocOrigin.Base, DocOrigin.Theirs });
        docMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Mine, DocOrigin.Theirs });
        docMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Mine, DocOrigin.Theirs, DocOrigin.Base });

        docMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Base, DocOrigin.Mine });
        docMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Base, DocOrigin.Mine, DocOrigin.Theirs });
        docMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Base, DocOrigin.Theirs });
        docMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Base, DocOrigin.Theirs, DocOrigin.Mine });

        docMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Theirs, DocOrigin.Base });
        docMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Theirs, DocOrigin.Base, DocOrigin.Mine });
        docMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Theirs, DocOrigin.Mine });
        docMergeOrderCandidates.Add(new List<DocOrigin>() { DocOrigin.Theirs, DocOrigin.Mine, DocOrigin.Base });

        // exclude
        var docOriginsToExclude = baseHunkInfo.ToExcludeDocOrigins();
        foreach (var docOrigin in docOriginsToExclude)
            docMergeOrderCandidates.RemoveAll(r => r != null && r.Contains(docOrigin));

        DocMergeOrderCandidates = docMergeOrderCandidates;
    }

    public bool IsConflict
    {
        get => DocMergeOrder is null;
    }
}
