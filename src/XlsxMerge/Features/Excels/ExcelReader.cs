using System.Runtime.InteropServices;
using XlsxMerge.Extensions;
using Excel = Microsoft.Office.Interop.Excel;

namespace XlsxMerge.Features.Excels;

public class ExcelReader : IDisposable
{
    private Excel.Application? _application;

    public ExcelReader()
    {
        _application = null;
    }

    public ExcelFile Read(string filePath)
    {
        if (_application == null)
        {
            // "xlsx 파일 비교 [3단계 중 1단계]", "엑셀을 실행하고 있습니다."
            _application = new Excel.Application();
            _application.DisplayAlerts = false;
            _application.Visible = false;
        }

        // "xlsx 파일 비교 [3단계 중 2단계]", "문서를 읽고 있습니다."
        string fullPath = Path.GetFullPath(filePath);
        var workbooks = _application.Workbooks;
        var workbook = workbooks.Open(fullPath);

        // 워크시트 순회하면서 셀 정보 읽기
        var excelFile = new ExcelFile();
        var worksheets = workbook!.Worksheets.GetWorksheets();
        foreach (var worksheet in worksheets)
        {
            if (worksheet == null)
                continue;

            worksheet.Activate();

            var excelWorksheet = ExcelWorksheet.Of(worksheet.Name);
            excelFile.Worksheets.Add(excelWorksheet);

            int rowCount = worksheet.GetRowCount();
            int colCount = worksheet.GetColumnCount();
            if (rowCount == 0 || colCount == 0)
                continue;

            var allCells = worksheet.GetAllCells(rowCount, colCount);

            // 컬럼 너비 읽기
            var columnsWidth = allCells.GetColumnsWidth(colCount);
            excelWorksheet.ColumnWidthList.AddRange(columnsWidth);

            // 셀 데이터 읽기
            var cells = allCells.GetCells(rowCount, colCount);
            excelWorksheet.Cells.AddRange(cells);

            Marshal.ReleaseComObject(allCells);
            allCells = null;

            Marshal.ReleaseComObject(worksheet);
        }

        workbook.Close(SaveChanges: false);  // https://stackoverflow.com/a/8813945

        Marshal.ReleaseComObject(workbooks);
        Marshal.ReleaseComObject(workbook);

        return excelFile;
    }

    public void Dispose()
    {
        if (_application != null)
        {
            _application.Quit();
            Marshal.ReleaseComObject(_application);
            _application = null;
            GC.Collect();
        }
    }
}
