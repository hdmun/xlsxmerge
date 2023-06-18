using XlsxMerge.Diff;
using XlsxMerge.Features.Diffs;

namespace XlsxMerge.Merge;

public class SheetMergeDecision
{
    public string WorksheetName;
    public readonly SheetDiffResult SheetDiffResult;
    public WorksheetMergeMode MergeModeDecision;
    public List<WorksheetMergeMode> MergeModeCandidates;
    public List<HunkMergeDecision> HunkMergeDecisionList;

    public SheetMergeDecision(SheetDiffResult sheetDiffResult)
    {
        SheetDiffResult = sheetDiffResult;
        WorksheetName = SheetDiffResult.WorksheetName;

        MergeModeCandidates = BuildMergeModeCandidates();
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

    private List<WorksheetMergeMode> BuildMergeModeCandidates()
    {
        var candidateList = new List<WorksheetMergeMode>();

        if (SheetDiffResult.ComparisonMode == ComparisonMode.TwoWay)
        {
            // Two-way merge
            // a1. Base있음 + Mine동일 = Unchanged
            // a2. Base있음 + Mine변경 = Merge
            // a3. Base있음 + Mine없음 = Delete, UseBase
            // a4. Base없음 + Mine생성 = UseMine

            var docsContaining = SheetDiffResult.DocsContaining;
            if (docsContaining.Contains(DocOrigin.Base))
            {
                if (docsContaining.Contains(DocOrigin.Mine))
                {
                    if (SheetDiffResult.HunkList.Find(r => r.hunkStatus == Diff3HunkStatus.MineDiffers) == null)
                        candidateList.Add(WorksheetMergeMode.Unchanged); // a1
                    else
                        candidateList.Add(WorksheetMergeMode.Merge); // a2
                }
                else
                {
                    candidateList.Add(WorksheetMergeMode.Delete); // a3
                    candidateList.Add(WorksheetMergeMode.UseBase); // a3

                }
            }
            else
            {
                candidateList.Add(WorksheetMergeMode.UseMine); // a4
            }
        }

        if (SheetDiffResult.ComparisonMode == ComparisonMode.ThreeWay)
        {
            // Three-way merge
            // Base있음 + Mine동일 + Theirs동일 = Unchanged

            // Base있음 + Mine동일 + Theirs변경 = Merge
            // Base있음 + Mine변경 + Theirs동일 = Merge
            // Base있음 + Mine변경 + Theirs변경 = Merge

            // Base있음 + Mine변경 + Theirs없음 = UseMine, UseBase, Delete
            // Base있음 + Mine없음 + Theirs변경 = UseTheirs, UseBase, Delete

            // Base있음 + Mine동일 + Theirs없음 = Delete, UseBase
            // Base있음 + Mine없음 + Theirs동일 = Delete, UseBase
            // Base있음 + Mine없음 + Theirs없음 = Delete, UseBase

            // Base없음 + Mine생성 + Theirs없음 = UseMine
            // Base없음 + Mine없음 + Theirs생성 = UseTheirs
            // Base없음 + Mine생성 + Theirs생성 = Merge

            var docsContaining = SheetDiffResult.DocsContaining;
            if (docsContaining.Contains(DocOrigin.Base))
            {
                if (SheetDiffResult.HunkList.Count == 0)
                {
                    candidateList.Add(WorksheetMergeMode.Unchanged);
                }
                else if (docsContaining.Contains(DocOrigin.Mine) && docsContaining.Contains(DocOrigin.Theirs))
                {
                    candidateList.Add(WorksheetMergeMode.Merge);
                    if (SheetDiffResult.HunkList.Find(r => r.hunkStatus == Diff3HunkStatus.MineDiffers) != null)
                        candidateList.Add(WorksheetMergeMode.UseMine);
                    if (SheetDiffResult.HunkList.Find(r => r.hunkStatus == Diff3HunkStatus.TheirsDiffers) != null)
                        candidateList.Add(WorksheetMergeMode.UseTheirs);
                    candidateList.Add(WorksheetMergeMode.UseBase);
                }
                else
                {
                    // 둘 중에 하나가 없는 상태.
                    if (SheetDiffResult.HunkList.Find(r => r.hunkStatus == Diff3HunkStatus.Conflict) != null)
                    {
                        if (docsContaining.Contains(DocOrigin.Mine))
                            candidateList.Add(WorksheetMergeMode.UseMine);
                        if (docsContaining.Contains(DocOrigin.Theirs))
                            candidateList.Add(WorksheetMergeMode.UseTheirs);
                        candidateList.Add(WorksheetMergeMode.UseBase);
                        candidateList.Add(WorksheetMergeMode.Delete);
                    }
                    else
                    {
                        candidateList.Add(WorksheetMergeMode.Delete);
                        candidateList.Add(WorksheetMergeMode.UseBase);
                    }
                }
            }
            else
            {
                if (docsContaining.Contains(DocOrigin.Mine) && docsContaining.Contains(DocOrigin.Theirs) == false)
                {
                    candidateList.Add(WorksheetMergeMode.UseMine);
                    candidateList.Add(WorksheetMergeMode.Delete);
                }
                if (docsContaining.Contains(DocOrigin.Mine) == false && docsContaining.Contains(DocOrigin.Theirs))
                {
                    candidateList.Add(WorksheetMergeMode.UseTheirs);
                    candidateList.Add(WorksheetMergeMode.Delete);
                }
                if (docsContaining.Contains(DocOrigin.Mine) && docsContaining.Contains(DocOrigin.Theirs))
                {
                    candidateList.Add(WorksheetMergeMode.Merge);
                    candidateList.Add(WorksheetMergeMode.UseMine);
                    candidateList.Add(WorksheetMergeMode.UseTheirs);
                    candidateList.Add(WorksheetMergeMode.Delete);
                }
            }
        }

        return candidateList;
    }
}
