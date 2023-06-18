using XlsxMerge.Features.Diffs;

namespace XlsxMerge.Diff;

public class SheetDiffResult
{
    public readonly string WorksheetName;
    public readonly ComparisonMode ComparisonMode;
    public readonly List<DocOrigin> DocsContaining; // 이 워크시트가 있는 문서.
    public List<DiffHunkInfo> HunkList = new();

    public SheetDiffResult(string worksheetName, ComparisonMode comparisonMode)
    {
        WorksheetName = worksheetName;
        ComparisonMode = comparisonMode;
        DocsContaining = new();
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
