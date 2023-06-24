using Newtonsoft.Json;

namespace XlsxMerge.Features.Excels;

public class ExcelFile
{
    public readonly List<ExcelWorksheet> Worksheets = new();

    public string[]? GetTextLinesByWorksheetName(string worksheetName)
    {
        var targetWorksheet = Worksheets.Find(r => r.Name == worksheetName);
        if (targetWorksheet == null)
            return null;

        var textByLines = new List<string>();
        foreach (var eachRow in targetWorksheet.Cells)
        {
            // 빈 컬럼 제거
            var columnList = eachRow.Select(r => r.ContentsForDiff3).ToList();
            while (columnList.Count > 0 && columnList[columnList.Count - 1] == "")
                columnList.RemoveAt(columnList.Count - 1);

            textByLines.Add(JsonConvert.SerializeObject(columnList, Formatting.None));
        }
        return textByLines.ToArray();
    }
}
