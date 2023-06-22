using System.Collections.Immutable;
using XlsxMerge.Diff;
using XlsxMerge.Features;
using XlsxMerge.Features.Diffs;
using XlsxMerge.Features.Excels;
using XlsxMerge.ViewModel;

namespace XlsxMerge
{
    class XlsxDiff3Core
    {
        private readonly Dictionary<DocOrigin, ExcelFile> ParsedWorkbookMap = new();

        public List<SheetDiffResult> Run(PathViewModel pathViewModel)
        {
            ReadExcelFiles(pathViewModel.BasePath, pathViewModel.MinePath, pathViewModel.TheirsPath, pathViewModel.ComparisonMode);

            FakeBackgroundWorker.OnUpdateProgress("xlsx 파일 비교  [3단계 중 3단계]", "엑셀 문서 비교 중..");

            // 비교 대상 워크시트 목록을 추출
            var sheetNameSet = ParsedWorkbookMap.Values
                .SelectMany(x => x.Worksheets.Select(y => y.Name))
                .ToHashSet();

            var docOriginEnums = Enum.GetValues<DocOrigin>();

            // 각 워크시트를 List<String>으로 변환 후 do diff3
            var compareResults = new List<SheetDiffResult>();
            foreach (var worksheetName in sheetNameSet)
            {
                var textLinesByOrigin = docOriginEnums.ToDictionary(
                    x => x,
                    x => ParsedWorkbookMap[x].GetTextLinesByWorksheetName(worksheetName)
                );

                var containDocs = textLinesByOrigin.Where(x => x.Value != null)
                    .Select(x => x.Key)
                    .ToImmutableHashSet();

                var baseTextLines = textLinesByOrigin[DocOrigin.Base];
                var mineTextLines = textLinesByOrigin[DocOrigin.Mine];
                var theirsTextLines = textLinesByOrigin[DocOrigin.Theirs];
                string diff3ResultText = LaunchExternalDiff3Process(baseTextLines, mineTextLines, theirsTextLines);

                var diff3Parser = new Diff3Parser();
                var parsedHunkList = diff3Parser.Parse(diff3ResultText);

                var newSheetResult = SheetDiffResult.Of(worksheetName, pathViewModel.ComparisonMode, containDocs, parsedHunkList);
                compareResults.Add(newSheetResult);
            }
            return compareResults;
        }

        private void ReadExcelFiles(string basePath, string minePath, string theirsPath, ComparisonMode comparison)
        {
            // 엑셀 파일을 해석.
            // diff3 의 file1/file2/file3 셋업.
            // 순서는 base/mine/theirs 이며, two-way는 theirs 위치에 base 문서를 넣는다.
            // 이렇게 하면 two-way에서 diff3 hunk status 가 Diff3HunkStatus.MineDiffers 로 표시된다.
            using (var excelReader = new ExcelReader())
            {
                ParsedWorkbookMap[DocOrigin.Base] = excelReader.Read(basePath);
                ParsedWorkbookMap[DocOrigin.Mine] = excelReader.Read(minePath);
                if (comparison == ComparisonMode.ThreeWay)
                    ParsedWorkbookMap[DocOrigin.Theirs] = excelReader.Read(theirsPath);
                else
                    ParsedWorkbookMap[DocOrigin.Theirs] = ParsedWorkbookMap[DocOrigin.Base];
            }
        }

	    public Dictionary<DocOrigin, ExcelWorksheet?> GetParsedWorksheetData(string worksheetName)
	    {
            return ParsedWorkbookMap.ToDictionary(
                x => x.Key,
                x => x.Value.Worksheets
                    .FirstOrDefault(y => y.Name == worksheetName)
            );
	    }

	    private static string LaunchExternalDiff3Process(List<string> lines1, List<string> lines2, List<string> lines3)
	    {
		    string tmp1 = Diff3Process.CreateTempFile(lines1.ToArray());
		    string tmp2 = Diff3Process.CreateTempFile(lines2.ToArray());
            string tmp3 = Diff3Process.CreateTempFile(lines3.ToArray());

            var diffFiles = new string[] { tmp1, tmp2, tmp3 };
		    return Diff3Process.Start(diffFiles);
	    }
    }
}
