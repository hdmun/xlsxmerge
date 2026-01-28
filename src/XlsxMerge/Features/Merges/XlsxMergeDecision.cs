using XlsxMerge.Features.Diffs;

namespace XlsxMerge.Features.Merges
{
    public class XlsxMergeDecision
    {
        public readonly List<SheetMergeDecision> SheetMergeDecisionList;

        public XlsxMergeDecision(List<SheetDiffResult> compareResults)
        {
            SheetMergeDecisionList = new();
            foreach (var result in compareResults)
                SheetMergeDecisionList.Add(new SheetMergeDecision(result));
        }

        public int CalcUnResolvedConflictCount()
        {
            return SheetMergeDecisionList
                .Where(x => x.MergeModeDecision == WorksheetMergeMode.Merge)
                .Sum(x => x.GetConflictCount());
        }
    }
}
