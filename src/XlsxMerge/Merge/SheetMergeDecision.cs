using XlsxMerge.Diff;
using XlsxMerge.Features.Diffs;

namespace XlsxMerge.Merge;

public class SheetMergeDecision
{
    public readonly string WorksheetName;
    public readonly SheetDiffResult SheetDiffResult;
    public WorksheetMergeMode MergeModeDecision;
    public readonly List<WorksheetMergeMode> MergeModeCandidates;
    public readonly List<HunkMergeDecision> HunkMergeDecisionList;

    public SheetMergeDecision(SheetDiffResult sheetDiffResult)
    {
        SheetDiffResult = sheetDiffResult;
        WorksheetName = SheetDiffResult.WorksheetName;

        MergeModeCandidates = SheetDiffResult.ToMergeMode();
        MergeModeDecision = MergeModeCandidates[0];

        HunkMergeDecisionList = new List<HunkMergeDecision>();
        foreach (var hunkInfo in SheetDiffResult.HunkList)
            HunkMergeDecisionList.Add(new HunkMergeDecision(hunkInfo));

        for (int i = 0; i < HunkMergeDecisionList.Count; i++)
        {
            var hunkInfo = SheetDiffResult.HunkList[i];

            // base에 비해 mine/theirs 모두 변경사항이 있고, 그 둘이 같으면 Mine을 사용한다.
            if (hunkInfo.hunkStatus == Diff3HunkStatus.BaseDiffers || hunkInfo.hunkStatus == Diff3HunkStatus.MineDiffers)
                HunkMergeDecisionList[i].DocMergeOrder.Add(DocOrigin.Mine);

            if (hunkInfo.hunkStatus == Diff3HunkStatus.TheirsDiffers)
                HunkMergeDecisionList[i].DocMergeOrder.Add(DocOrigin.Theirs);

            if (hunkInfo.hunkStatus == Diff3HunkStatus.Conflict)
                HunkMergeDecisionList[i].DocMergeOrder = null;
        }
    }
}
