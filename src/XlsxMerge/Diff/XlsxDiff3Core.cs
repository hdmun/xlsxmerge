using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using XlsxMerge.Diff;
using XlsxMerge.ViewModel;
using XlsxMerge.Features;

namespace XlsxMerge
{
    class XlsxDiff3Core
    {
        public class SheetDiffResult
        {
	        public ComparisonMode ComparisonMode = ComparisonMode.Unknown;
			public string WorksheetName = "";
            public List<DocOrigin> DocsContaining = new List<DocOrigin>(); // 이 워크시트가 있는 문서.
            public List<DiffHunkInfo> HunkList = new List<DiffHunkInfo>();

            public class DiffHunkInfo
            {
                public Diff3HunkStatus hunkStatus;
                public Dictionary<DocOrigin, RowRange> rowRangeMap = new Dictionary<DocOrigin, RowRange>();
            }
        }

        public Dictionary<DocOrigin, ExcelFile> ParsedWorkbookMap = new();

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
			var xlsxList = new List<ExcelFile>();
            xlsxList.Add(ParsedWorkbookMap[DocOrigin.Base]);
            xlsxList.Add(ParsedWorkbookMap[DocOrigin.Mine]);
            xlsxList.Add(pathViewModel.ComparisonMode == ComparisonMode.ThreeWay ? ParsedWorkbookMap[DocOrigin.Theirs] : ParsedWorkbookMap[DocOrigin.Base]);

            // 비교 대상 워크시트 목록을 추출
            List<String> allSheetNameList = new List<string>();
            foreach (var eachXlsx in xlsxList)
                foreach (string sheetName in eachXlsx.Worksheets.Select(r => r.Name))
                    if (allSheetNameList.Contains(sheetName) == false)
                        allSheetNameList.Add(sheetName);

            // 각 워크시트를 List<String>으로 변환 후 do diff3
            var compareResults = new List<SheetDiffResult>();
            foreach (var worksheetName in allSheetNameList)
            {
                SheetDiffResult newSheetResult = new SheetDiffResult();
                newSheetResult.WorksheetName = worksheetName;
	            newSheetResult.ComparisonMode = pathViewModel.ComparisonMode;

				string diff3ResultText = null;
                {
                    var lines1 = getWorksheetLines(xlsxList[0], worksheetName);
                    var lines2 = getWorksheetLines(xlsxList[1], worksheetName);
                    var lines3 = getWorksheetLines(xlsxList[2], worksheetName);

                    if (lines1 != null)
                        newSheetResult.DocsContaining.Add(DocOrigin.Base);
                    if (lines2 != null)
                        newSheetResult.DocsContaining.Add(DocOrigin.Mine);
                    if (lines3 != null)
                        newSheetResult.DocsContaining.Add(DocOrigin.Theirs);

                    diff3ResultText = LaunchExternalDiff3Process(lines1, lines2, lines3);
                }

                newSheetResult.HunkList = parseDiff3Result(diff3ResultText);

                compareResults.Add(newSheetResult);
            }
            return compareResults;
        }

	    public Dictionary<DocOrigin, ExcelWorksheet> GetParsedWorksheetData(string worksheetName)
	    {
		    var result = new Dictionary<DocOrigin, ExcelWorksheet>();
		    foreach (var docOrigin in new[] {DocOrigin.Base, DocOrigin.Mine, DocOrigin.Theirs})
		    {
			    result[docOrigin] = null;
			    if (ParsedWorkbookMap.ContainsKey(docOrigin))
				    result[docOrigin] = ParsedWorkbookMap[docOrigin].Worksheets.FirstOrDefault(r => r.Name == worksheetName);
			}

		    return result;
	    }

		private static List<String> getWorksheetLines(ExcelFile xlsxFile, String worksheetName)
        {
            var targetWorksheet = xlsxFile.Worksheets.Find(r => r.Name == worksheetName);
            if (targetWorksheet == null)
                return null;

            List<String> result = new List<string>();
            foreach (var eachRow in targetWorksheet.Cells)
            {
                var columnList = eachRow.Select(r => r.ContentsForDiff3).ToList();
                while (columnList.Count > 0 && columnList[columnList.Count - 1] == "")
                    columnList.RemoveAt(columnList.Count - 1);
                result.Add(JsonConvert.SerializeObject(columnList, Newtonsoft.Json.Formatting.None));
            }
            return result;
        }

        private static List<SheetDiffResult.DiffHunkInfo> parseDiff3Result(String diff3ResultText)
        {
            var regexLineInfo = new Regex("^([123]):([0-9,]+)([ac])$");
            var hunkStatusMap = new Dictionary<string, Diff3HunkStatus>()
            {
                { "====", Diff3HunkStatus.Conflict},
                { "====1", Diff3HunkStatus.BaseDiffers},
                { "====2", Diff3HunkStatus.MineDiffers},
                { "====3", Diff3HunkStatus.TheirsDiffers}
            };
            var FileOrderMap = new Dictionary<string, DocOrigin>()
            {
                { "1", DocOrigin.Base },
                { "2", DocOrigin.Mine },
                { "3", DocOrigin.Theirs },
            };

            var hunkInfoList = new List<SheetDiffResult.DiffHunkInfo>();

            SheetDiffResult.DiffHunkInfo curHunk = null;
            StringReader sr = new StringReader(diff3ResultText);
            while (sr.Peek() != -1)
            {
                var curLine = sr.ReadLine();
                if (curLine.StartsWith("===="))
                {
                    curHunk = new SheetDiffResult.DiffHunkInfo();
                    hunkInfoList.Add(curHunk);
                    curHunk.hunkStatus = hunkStatusMap[curLine.Trim()];
                }

                Match m = regexLineInfo.Match(curLine);
                if (m.Success == false)
                    continue;

                string fileIndex = m.Groups[1].Value;
                string[] rangeToken = m.Groups[2].Value.Split(new char[] { ',' });
                string command = m.Groups[3].Value;

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
                curHunk.rowRangeMap[rowRangeDoc] = rowRangeValue;
            }
            sr.Dispose();

            return hunkInfoList;
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
