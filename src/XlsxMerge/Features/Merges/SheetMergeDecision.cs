using XlsxMerge.Features.Diffs;

namespace XlsxMerge.Features.Merges;

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
        WorksheetName = sheetDiffResult.WorksheetName;

        MergeModeCandidates = sheetDiffResult.ToMergeMode();
        MergeModeDecision = MergeModeCandidates[0];

        HunkMergeDecisionList = new List<HunkMergeDecision>();
        foreach (var hunkInfo in sheetDiffResult.HunkList)
            HunkMergeDecisionList.Add(new HunkMergeDecision(hunkInfo));
    }

    public int GetConflictCount()
    {
        return HunkMergeDecisionList.Count(x => x.IsConflict);
    }
}
