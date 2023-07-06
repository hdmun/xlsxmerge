namespace XlsxMerge.Features.Excels;

public class ExcelCellContents
{
    public readonly string Value2String;
    public readonly string FormulaString;
    public readonly string ContentsForDiff3; // '같은 값' 비교에 사용합니다.

    public ExcelCellContents(string _value2String = "", string _formulaString = "")
    {
        Value2String = _value2String;
        FormulaString = _formulaString;
        if (FormulaString == null || FormulaString.StartsWith("=") == false)
            FormulaString = "";
        ContentsForDiff3 = string.IsNullOrWhiteSpace(FormulaString) ? Value2String : FormulaString;
    }
}
