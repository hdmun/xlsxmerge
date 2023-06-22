using System.Text.RegularExpressions;
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
			// 엑셀 파일을 해석.
			using (var excelReader = new ExcelReader())
            {
                ParsedWorkbookMap[DocOrigin.Base] = excelReader.Read(pathViewModel.BasePath);
                ParsedWorkbookMap[DocOrigin.Mine] = excelReader.Read(pathViewModel.MinePath);
                if (pathViewModel.ComparisonMode == ComparisonMode.ThreeWay)
                    ParsedWorkbookMap[DocOrigin.Theirs] = excelReader.Read(pathViewModel.TheirsPath);
            }

	        FakeBackgroundWorker.OnUpdateProgress("xlsx 파일 비교  [3단계 중 3단계]", "엑셀 문서 비교 중..");
            // diff3 의 file1/file2/file3 셋업.
            // 순서는 base/mine/theirs 이며, two-way는 theirs 위치에 base 문서를 넣는다.
            // 이렇게 하면 two-way에서 diff3 hunk status 가 Diff3HunkStatus.MineDiffers 로 표시된다.

            var xlsxList = new List<ExcelFile>
            {
                ParsedWorkbookMap[DocOrigin.Base],
                ParsedWorkbookMap[DocOrigin.Mine],
                pathViewModel.ComparisonMode switch
                {
                    ComparisonMode.ThreeWay => ParsedWorkbookMap[DocOrigin.Theirs],
                    _ => ParsedWorkbookMap[DocOrigin.Base]
                }
            };

            // 비교 대상 워크시트 목록을 추출
            var sheetNameSet = xlsxList
                .SelectMany(x => x.Worksheets.Select(y => y.Name))
                .ToHashSet();

            // 각 워크시트를 List<String>으로 변환 후 do diff3
            var compareResults = new List<SheetDiffResult>();
            foreach (var worksheetName in sheetNameSet)
            {
                SheetDiffResult newSheetResult = new SheetDiffResult(worksheetName, pathViewModel.ComparisonMode);

				string diff3ResultText = null;
                {
                    var lines1 = xlsxList[0].GetTextLinesByWorksheetName(worksheetName);
                    var lines2 = xlsxList[1].GetTextLinesByWorksheetName(worksheetName);
                    var lines3 = xlsxList[2].GetTextLinesByWorksheetName(worksheetName);

                    if (lines1 != null)
                        newSheetResult.DocsContaining.Add(DocOrigin.Base);
                    if (lines2 != null)
                        newSheetResult.DocsContaining.Add(DocOrigin.Mine);
                    if (lines3 != null)
                        newSheetResult.DocsContaining.Add(DocOrigin.Theirs);

                    diff3ResultText = LaunchExternalDiff3Process(lines1, lines2, lines3);
                }

                newSheetResult.HunkList = new Diff3Parser().Parse(diff3ResultText);

                compareResults.Add(newSheetResult);
            }
            return compareResults;
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
