using XlsxMerge.Diff;
using XlsxMerge.Merge;

namespace XlsxMerge
{
    class XlsxMergeDecision
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
            int unResolvedConflictCount = 0;
            foreach (var mergeDecision in SheetMergeDecisionList)
            {
                if (mergeDecision.MergeModeDecision != WorksheetMergeMode.Merge)
                    continue;

                unResolvedConflictCount += mergeDecision.HunkMergeDecisionList.Count(x => x == null);
            }
            return unResolvedConflictCount;
        }
    }
}
