using System.Collections.Immutable;
using XlsxMerge.Features.Diffs;
using XlsxMerge.Model;

namespace XlsxMerge.Diff;

public class SheetDiffResult
{
    public static SheetDiffResult Of(string worksheetName, ComparisonMode comparisonMode, ImmutableHashSet<DocOrigin> docsContaining, List<DiffHunkInfo> hunkList)
    {
        return new SheetDiffResult(worksheetName, comparisonMode, docsContaining, hunkList);
    }

    public readonly string WorksheetName;
    public readonly ComparisonMode ComparisonMode;
    private readonly ImmutableHashSet<DocOrigin> _containDocs; // 이 워크시트가 있는 문서.
    public readonly List<DiffHunkInfo> HunkList;

    private SheetDiffResult(string worksheetName, ComparisonMode comparisonMode, ImmutableHashSet<DocOrigin> docsContaining, List<DiffHunkInfo> hunkList)
    {
        WorksheetName = worksheetName;
        ComparisonMode = comparisonMode;
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
}
