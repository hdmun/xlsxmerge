using System.Collections.Immutable;
using XlsxMerge.Features.Diffs;

namespace XlsxMerge.Diff;

public class SheetDiffResult
{
    public static SheetDiffResult Of(string worksheetName, ComparisonMode comparisonMode, ImmutableHashSet<DocOrigin> docsContaining, List<DiffHunkInfo> hunkList)
    {
        return new SheetDiffResult(worksheetName, comparisonMode, docsContaining, hunkList);
    }

    public readonly string WorksheetName;
    public readonly ComparisonMode ComparisonMode;
    public readonly ImmutableHashSet<DocOrigin> DocsContaining; // 이 워크시트가 있는 문서.
    public readonly List<DiffHunkInfo> HunkList;

    private SheetDiffResult(string worksheetName, ComparisonMode comparisonMode, ImmutableHashSet<DocOrigin> docsContaining, List<DiffHunkInfo> hunkList)
    {
        WorksheetName = worksheetName;
        ComparisonMode = comparisonMode;
        DocsContaining = docsContaining;
        HunkList = hunkList;
    }

    public ModificationStateModel GetModificationSummary(DocOrigin targetDoc)
    {
        if (DocsContaining.Contains(DocOrigin.Base) == true && DocsContaining.Contains(targetDoc) == false)
            return new ModificationStateModel("삭제됨", Color.PaleVioletRed);

        if (DocsContaining.Contains(DocOrigin.Base) == false && DocsContaining.Contains(targetDoc) == true)
            return new ModificationStateModel("추가됨", Color.PaleGreen);

        if (HunkList.Find(r => r.hunkStatus == Diff3HunkStatus.Conflict) != null)
            return new ModificationStateModel("수정됨", Color.LightYellow);

        if (HunkList.Find(r => r.hunkStatus == Diff3HunkStatus.BaseDiffers) != null)
            return new ModificationStateModel("수정됨", Color.LightYellow);

        Diff3HunkStatus targetDocDiffers = Diff3HunkStatus.Conflict;
        if (targetDoc == DocOrigin.Mine)
            targetDocDiffers = Diff3HunkStatus.MineDiffers;
        if (targetDoc == DocOrigin.Theirs)
            targetDocDiffers = Diff3HunkStatus.TheirsDiffers;

        if (HunkList.Find(r => r.hunkStatus == targetDocDiffers) != null)
            return new ModificationStateModel("수정됨", Color.LightYellow);
        return new ModificationStateModel("같음", Color.White);
    }
}
