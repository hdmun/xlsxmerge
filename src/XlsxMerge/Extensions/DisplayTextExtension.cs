namespace XlsxMerge.Extensions;

public static class DisplayTextExtension
{
    public static string GetDisplayText(this WorksheetMergeMode mergeMode)
    {
        return mergeMode switch
        {
            WorksheetMergeMode.Unchanged => "변경 사항 없음",
            WorksheetMergeMode.Delete => "이 워크시트 삭제",
            WorksheetMergeMode.UseBase => "Base 버전으로 워크시트 사용",
            WorksheetMergeMode.UseMine => "Mine 버전으로 워크시트 사용",
            WorksheetMergeMode.UseTheirs => "Theirs 버전으로 워크시트 사용",
            WorksheetMergeMode.Merge => "변경점 직접 머지",
            _ => ""
        }; ;
    }

    public static string GetDisplayText(this List<DocOrigin> candidate)
    {
        if (candidate == null)
            return "충돌 상태 그대로 두기";
        if (candidate.Count == 0)
            return "삭제하기";
        if (candidate.Count == 1)
            return $"{candidate[0]} 변경점만 적용";
        return $"조합 : {string.Join(" > ", candidate.Select(r => r.ToString()))}";
    }
}
