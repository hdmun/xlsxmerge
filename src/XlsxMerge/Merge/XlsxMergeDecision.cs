using XlsxMerge.Diff;
using XlsxMerge.Merge;
using XlsxMerge.ViewModel;

namespace XlsxMerge
{
    class XlsxMergeDecision
    {
        public readonly List<SheetMergeDecision> SheetMergeDecisionList;
        public readonly DiffViewModel DiffViewModel;

        public XlsxMergeDecision(DiffViewModel diffViewModel, List<SheetDiffResult> compareResults)
        {
            DiffViewModel = diffViewModel;
            SheetMergeDecisionList = new List<SheetMergeDecision>();

            foreach (var result in compareResults)
                SheetMergeDecisionList.Add(new SheetMergeDecision(result));
        }
    }
}
