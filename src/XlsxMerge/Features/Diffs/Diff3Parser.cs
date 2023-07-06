using System.Text.RegularExpressions;

namespace XlsxMerge.Features.Diffs;

public class Diff3Parser
{
    private static readonly Regex _regexLineInfo = new("^([123]):([0-9,]+)([ac])$");

    private static readonly Dictionary<string, Diff3HunkStatus> _hunkStatusMap = new()
    {
        { "====", Diff3HunkStatus.Conflict},
        { "====1", Diff3HunkStatus.BaseDiffers},
        { "====2", Diff3HunkStatus.MineDiffers},
        { "====3", Diff3HunkStatus.TheirsDiffers}
    };

    private static readonly Dictionary<string, DocOrigin> FileOrderMap = new()
    {
        { "1", DocOrigin.Base },
        { "2", DocOrigin.Mine },
        { "3", DocOrigin.Theirs },
    };
    
    private readonly List<DiffHunkInfo> _hunkInfoList;

    public Diff3Parser()
    {
        _hunkInfoList = new();
    }

    public List<DiffHunkInfo> Parse(string diff3ResultText)
    {
        _hunkInfoList.Clear();
        DiffHunkInfo? hunkInfo = null;

        StringReader sr = new StringReader(diff3ResultText);
        while (sr.Peek() != -1)
        {
            var curLine = sr.ReadLine();
            if (curLine.StartsWith("===="))
            {
                var text = curLine.Trim();
                var hunStatus = _hunkStatusMap[text];
                hunkInfo = new DiffHunkInfo(hunStatus);
                _hunkInfoList.Add(hunkInfo);
                continue;
            }

            Match m = _regexLineInfo.Match(curLine);
            if (m.Success == false)
                continue;

            string fileIndex = m.Groups[1].Value;
            string[] rangeToken = m.Groups[2].Value.Split(new char[] { ',' });
            string command = m.Groups[3].Value;

            var (origin, rowRagne) = ParseInternal(fileIndex, rangeToken, command);
            hunkInfo.rowRangeMap[origin] = rowRagne;
        }
        sr.Dispose();

        return _hunkInfoList;
    }

    private (DocOrigin origin, RowRange rowRagne) ParseInternal(string fileIndex, string[] rangeToken, string command)
    {
        RowRange rowRangeValue = new RowRange();
        rowRangeValue.RowNumber = int.Parse(rangeToken[0]);
        rowRangeValue.RowCount = 0;

        if (command == "c")
        {
            rowRangeValue.RowCount = 1;
            if (rangeToken.Length > 1)
                rowRangeValue.RowCount = int.Parse(rangeToken[1]) - rowRangeValue.RowNumber + 1;
        }
        if (command == "a")
        {
            // diffutils 설명에 따르면 : 'FILE:La' This hunk appears after line L of file FILE, and contains no lines in that file. 
            // 'After'이므로, RowNumber를 1 더해주고 RowCount를 0으로 한다.
            rowRangeValue.RowNumber = int.Parse(rangeToken[0]) + 1;
            rowRangeValue.RowCount = 0;
        }

        DocOrigin rowRangeDoc = FileOrderMap[fileIndex];
        return (rowRangeDoc, rowRangeValue);
    }
}
