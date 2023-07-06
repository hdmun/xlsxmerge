using System.Collections.Immutable;
using XlsxMerge.Model;

namespace XlsxMerge.Features.Diffs;

public class SheetDiffResult
{
    public static SheetDiffResult Of(string worksheetName, ComparisonMode comparisonMode, ImmutableHashSet<DocOrigin> docsContaining, List<DiffHunkInfo> hunkList)
    {
        return new SheetDiffResult(worksheetName, comparisonMode, docsContaining, hunkList);
    }

    public readonly string WorksheetName;
    private readonly ComparisonMode _comparisonMode;
    private readonly ImmutableHashSet<DocOrigin> _containDocs; // 이 워크시트가 있는 문서.
    public readonly List<DiffHunkInfo> HunkList;

    private SheetDiffResult(string worksheetName, ComparisonMode comparisonMode, ImmutableHashSet<DocOrigin> docsContaining, List<DiffHunkInfo> hunkList)
    {
        WorksheetName = worksheetName;
        _comparisonMode = comparisonMode;
        _containDocs = docsContaining;
        HunkList = hunkList;
    }

    public bool HasBaseDoc => HasDocOrigin(DocOrigin.Base);
    public bool HasMineDoc => HasDocOrigin(DocOrigin.Mine);
    public bool HasTheirsDoc => HasDocOrigin(DocOrigin.Theirs);

    public bool HasDocOrigin(DocOrigin docOrigin)
    {
        return _containDocs.Contains(docOrigin);
    }

    public bool IsEmtpyHunk => HunkList.Count == 0;

    public bool HasBaseDiffers => HasHunkStatus(Diff3HunkStatus.BaseDiffers);
    public bool HasMineDiffers => HasHunkStatus(Diff3HunkStatus.MineDiffers);
    public bool HasTheirsDiffers => HasHunkStatus(Diff3HunkStatus.TheirsDiffers);
    public bool HasConflict => HasHunkStatus(Diff3HunkStatus.Conflict);

    private bool HasHunkStatus(Diff3HunkStatus status)
    {
        return HunkList.Any(x => x.hunkStatus == status);
    }

    public ModificationStateModel GetModificationSummary(DocOrigin targetDoc)
    {
        if (HasBaseDoc && HasDocOrigin(targetDoc) == false)
            return new ModificationStateModel("삭제됨", Color.PaleVioletRed);

        if (!HasBaseDoc && HasDocOrigin(targetDoc) == true)
            return new ModificationStateModel("추가됨", Color.PaleGreen);

        if (HasConflict)
            return new ModificationStateModel("수정됨", Color.LightYellow);

        if (HasBaseDiffers)
            return new ModificationStateModel("수정됨", Color.LightYellow);

        Diff3HunkStatus targetDocDiffers = Diff3HunkStatus.Conflict;
        if (targetDoc == DocOrigin.Mine)
            targetDocDiffers = Diff3HunkStatus.MineDiffers;
        if (targetDoc == DocOrigin.Theirs)
            targetDocDiffers = Diff3HunkStatus.TheirsDiffers;

        if (HasHunkStatus(targetDocDiffers))
            return new ModificationStateModel("수정됨", Color.LightYellow);
        return new ModificationStateModel("같음", Color.White);
    }

    public List<WorksheetMergeMode> ToMergeMode()
    {
        switch (_comparisonMode)
        {
            case ComparisonMode.TwoWay:
                return ToTwoWayMergeMode();
            case ComparisonMode.ThreeWay:
                return ToThreeWayMergeMode();
            default:
                return new List<WorksheetMergeMode>();
        }
    }

    private List<WorksheetMergeMode> ToTwoWayMergeMode()
    {
        var candidateList = new List<WorksheetMergeMode>();

        // Two-way merge
        // a1. Base있음 + Mine동일 = Unchanged
        // a2. Base있음 + Mine변경 = Merge
        // a3. Base있음 + Mine없음 = Delete, UseBase
        // a4. Base없음 + Mine생성 = UseMine

        if (HasBaseDoc)
        {
            if (HasMineDoc)
            {
                if (HasMineDiffers == false)
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

        return candidateList;
    }

    private List<WorksheetMergeMode> ToThreeWayMergeMode()
    {
        var candidateList = new List<WorksheetMergeMode>();
        if (_comparisonMode != ComparisonMode.ThreeWay)
            return candidateList;

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

        if (HasBaseDoc)
        {
            if (IsEmtpyHunk)
            {
                candidateList.Add(WorksheetMergeMode.Unchanged);
            }
            else if (HasMineDoc && HasTheirsDoc)
            {
                candidateList.Add(WorksheetMergeMode.Merge);
                if (HasMineDiffers)
                    candidateList.Add(WorksheetMergeMode.UseMine);
                if (HasTheirsDiffers)
                    candidateList.Add(WorksheetMergeMode.UseTheirs);
                candidateList.Add(WorksheetMergeMode.UseBase);
            }
            else
            {
                // 둘 중에 하나가 없는 상태.
                if (HasConflict)
                {
                    if (HasMineDoc)
                        candidateList.Add(WorksheetMergeMode.UseMine);
                    if (HasTheirsDoc)
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
            if (HasMineDoc && !HasTheirsDoc)
            {
                candidateList.Add(WorksheetMergeMode.UseMine);
                candidateList.Add(WorksheetMergeMode.Delete);
            }

            if (!HasMineDoc && HasTheirsDoc)
            {
                candidateList.Add(WorksheetMergeMode.UseTheirs);
                candidateList.Add(WorksheetMergeMode.Delete);
            }

            if (HasMineDoc && HasTheirsDoc)
            {
                candidateList.Add(WorksheetMergeMode.Merge);
                candidateList.Add(WorksheetMergeMode.UseMine);
                candidateList.Add(WorksheetMergeMode.UseTheirs);
                candidateList.Add(WorksheetMergeMode.Delete);
            }
        }

        return candidateList;
    }
}
