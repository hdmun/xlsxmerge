using System.Runtime.InteropServices;
using XlsxMerge.Features;
using Excel = Microsoft.Office.Interop.Excel;

namespace XlsxMerge.Extensions;

public static class ExcelExtension
{
    public static IEnumerable<Excel.Worksheet?> GetWorksheets(this Excel.Sheets sheets)
    {
        for (int sheetIndex = 1; sheetIndex <= sheets.Count; sheetIndex++)
        {
            yield return sheets.Item[sheetIndex] as Excel.Worksheet;
        }
    }

    public static int GetRowCount(this Excel.Worksheet worksheet)
    {
        var range = worksheet.Find(Excel.XlSearchOrder.xlByRows);

        int rowCount = 0;
        if (range != null)
        {
            rowCount = range.Row;
            Marshal.ReleaseComObject(range);
        }
        return rowCount;
    }


    public static int GetColumnCount(this Excel.Worksheet worksheet)
    {
        var range = worksheet.Find(Excel.XlSearchOrder.xlByColumns);

        int colCount = 0;
        if (range != null)
        {
            colCount = range.Column;
            Marshal.ReleaseComObject(range);
        }

        return colCount;
    }

    public static Excel.Range Find(this Excel.Worksheet worksheet, Excel.XlSearchOrder xlSearchOrder)
    {
        var a1Cell = worksheet.Cells[1, 1] as Excel.Range;
        var wsCells = worksheet.Cells;

        var findRange = wsCells.Find("*", a1Cell,
            Excel.XlFindLookIn.xlFormulas,
            Excel.XlLookAt.xlPart,
            xlSearchOrder,
            Excel.XlSearchDirection.xlPrevious,
            false);

        Marshal.ReleaseComObject(wsCells);
        Marshal.ReleaseComObject(a1Cell);

        return findRange;
    }

    public static Excel.Range GetAllCells(this Excel.Worksheet worksheet, int rowCount, int colCount)
    {
        // endCell에는 행/열 번호에 1씩 더해줘서, Value2 / Value / Formula 가 single value / 0-based array로 동작하는 것을 막는다.
        // https://stackoverflow.com/a/37176162
        var startCell = worksheet.Cells[1, 1];
        var endCell = worksheet.Cells[rowCount + 1, colCount + 1];
        Excel.Range allCells = worksheet.Range[startCell.Address, endCell.Address];
        Marshal.ReleaseComObject(endCell);
        Marshal.ReleaseComObject(startCell);
        return allCells;
    }

    public static IEnumerable<double> GetColumnsWidth(this Excel.Range range, int colCount)
    {
        for (int c = 1; c <= colCount; c++)
        {
            var headerCell = range[1, c];
            var columnWidth = headerCell.Width;
            Marshal.ReleaseComObject(headerCell);

            yield return (double)columnWidth;
        }
    }

    public static IEnumerable<List<ExcelCellContents>> GetCells(this Excel.Range range, int rowCount, int colCount)
    {
        for (int row = 1; row <= rowCount; row++)
        {
            yield return range.GetRowCells(row, colCount).ToList();
        }
    }

    private static IEnumerable<ExcelCellContents> GetRowCells(this Excel.Range range, int row, int colCount)
    {
        // classic : https://fastexcel.wordpress.com/2011/11/30/text-vs-value-vs-value2-slow-text-and-how-to-avoid-it/
        var values = range.Value2 as object[,];
        var formula = range.FormulaR1C1 as object[,];

        for (int col = 1; col <= colCount; col++)
        {
            var newCell = new ExcelCellContents($"{values[row, col]}", $"{formula[row, col]}");
            yield return newCell;
        }
    }
}
