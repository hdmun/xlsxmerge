using XlsxMerge.Diff;
using XlsxMerge.Merge;

namespace XlsxMerge
{
    class XlsxMergeDecision
    {
        public readonly List<SheetMergeDecision> SheetMergeDecisionList;
        public readonly XlsxDiff3Core DiffResult;

        public XlsxMergeDecision(XlsxDiff3Core diffResult, List<SheetDiffResult> compareResults)
        {
            DiffResult = diffResult;
            SheetMergeDecisionList = new List<SheetMergeDecision>();

            if (diffResult == null)
                return;

            foreach (var result in compareResults)
                SheetMergeDecisionList.Add(new SheetMergeDecision(result));
        }
    }
}
