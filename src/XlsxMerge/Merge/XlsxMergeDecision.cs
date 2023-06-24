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
    }
}
