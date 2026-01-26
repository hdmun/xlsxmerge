namespace XlsxMerge.Features.Diffs.Enums;

public enum Diff3HunkStatus
{
    BaseDiffers, // base 문서만 다르다.
    MineDiffers, // mine 문서만 다르다.
    TheirsDiffers, // theirs 문서만 다르다.
    Conflict, // mine + theirs 각각 수정점이 있다.
}
