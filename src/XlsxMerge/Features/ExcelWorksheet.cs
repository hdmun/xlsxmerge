namespace XlsxMerge.Features;

public class ExcelWorksheet
{
    public static ExcelWorksheet Of(string name)
    {
        return new ExcelWorksheet(name);
    }

    public string Name { get; set; }
    public readonly List<double> ColumnWidthList;
    public readonly List<List<ExcelCellContents>> Cells;

    public ExcelWorksheet(string name)
    {
        Name = name;
        Cells = new(); // List<Row>
        ColumnWidthList = new();
    }

    public int RowCount
    {
        get => Cells.Count;
    }

    public int ColumnCount
    {
        get => Cells.Count == 0 ? 0 : Cells[0].Count;
    }

    public ExcelCellContents Cell(int rowNumber, int colNumber)
    {
        if (rowNumber <= Cells.Count && colNumber <= Cells[rowNumber - 1].Count)
            return Cells[rowNumber - 1][colNumber - 1];
        return new ExcelCellContents();
    }
}
