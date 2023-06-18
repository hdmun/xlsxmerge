namespace XlsxMerge.Features.Diffs;

public enum DocOrigin
{
    Base,
    Mine, // mine = destination = current
    Theirs, // theirs = source = others
}
