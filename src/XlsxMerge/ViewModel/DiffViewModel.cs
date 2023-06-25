using System;
using System.Collections.Immutable;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using XlsxMerge.Features.Diffs;
using XlsxMerge.Features.Excels;

namespace XlsxMerge.ViewModel;

public class DiffViewModel : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    private readonly Dictionary<DocOrigin, ExcelFile> _excelFiles;

    public DiffViewModel()
    {
        _excelFiles = new();
    }

    public void Read(ComparisonMode comparison, string basePath, string minePath, string theirsPath)
    {
        // 엑셀 파일을 해석.
        // diff3 의 file1/file2/file3 셋업.
        // 순서는 base/mine/theirs 이며, two-way는 theirs 위치에 base 문서를 넣는다.
        // 이렇게 하면 two-way에서 diff3 hunk status 가 Diff3HunkStatus.MineDiffers 로 표시된다.

        using (var excelReader = new ExcelReader())
        {
            _excelFiles[DocOrigin.Base] = excelReader.Read(basePath);
            _excelFiles[DocOrigin.Mine] = excelReader.Read(minePath);
            if (comparison == ComparisonMode.ThreeWay)
                _excelFiles[DocOrigin.Theirs] = excelReader.Read(theirsPath);
            else
                _excelFiles[DocOrigin.Theirs] = _excelFiles[DocOrigin.Base];
        }

        // "xlsx 파일 비교  [3단계 중 3단계]", "엑셀 문서 비교 중.."
    }

    public List<SheetDiffResult> DiffExcels(ComparisonMode comparison)
    {
        // 비교 대상 워크시트 목록을 추출
        var sheetNameSet = _excelFiles.Values
            .SelectMany(x => x.Worksheets.Select(y => y.Name))
            .ToHashSet();

        var docOriginEnums = Enum.GetValues<DocOrigin>();

        // 각 워크시트를 List<String>으로 변환 후 do diff3
        var compareResults = new List<SheetDiffResult>();
        foreach (var worksheetName in sheetNameSet)
        {
            var textLinesByOrigin = docOriginEnums.ToDictionary(
                x => x,
                x => _excelFiles[x].GetTextLinesByWorksheetName(worksheetName)
            );

            var containDocs = textLinesByOrigin.Where(x => x.Value != null)
                .Select(x => x.Key)
                .ToImmutableHashSet();

            var baseTextLines = textLinesByOrigin[DocOrigin.Base];
            var mineTextLines = textLinesByOrigin[DocOrigin.Mine];
            var theirsTextLines = textLinesByOrigin[DocOrigin.Theirs];
            string diff3ResultText = StartDiff3(baseTextLines, mineTextLines, theirsTextLines);

            var diff3Parser = new Diff3Parser();
            var parsedHunkList = diff3Parser.Parse(diff3ResultText);

            var newSheetResult = SheetDiffResult.Of(worksheetName, comparison, containDocs, parsedHunkList);
            compareResults.Add(newSheetResult);
        }
        return compareResults;
    }

    public Dictionary<DocOrigin, ExcelWorksheet?> GetWorksheets(string worksheetName)
    {
        return _excelFiles.ToDictionary(
            x => x.Key,
            x => x.Value.Worksheets
                .FirstOrDefault(y => y.Name == worksheetName)
        );
    }

    public ExcelWorksheet? GetWorksheetsBy(string worksheetName, DocOrigin docOrigin)
    {
        var worksheets = GetWorksheets(worksheetName);
        return worksheets.TryGetValue(docOrigin, out var worksheet) ? worksheet : null;
    }

    private string StartDiff3(string[]? baseTextByLines, string[]? mineTextByLines, string[]? thierTextByLines)
    {
        string tmp1 = Diff3Process.CreateTempFile(baseTextByLines);
        string tmp2 = Diff3Process.CreateTempFile(mineTextByLines);
        string tmp3 = Diff3Process.CreateTempFile(thierTextByLines);

        var diffFiles = new string[] { tmp1, tmp2, tmp3 };
        return Diff3Process.Start(diffFiles);
    }

    private void OnPropertyChanged([CallerMemberName] string propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
