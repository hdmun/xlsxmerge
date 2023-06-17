using System.ComponentModel;
using System.Runtime.CompilerServices;
using XlsxMerge.Diff;
using XlsxMerge.Features;

namespace XlsxMerge.ViewModel;

public class DiffViewModel : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    public IEnumerable<string> GetWorksheetColumns(ComparisonMode comparisonMode)
    {
        yield return "워크시트 이름";
        if (comparisonMode == ComparisonMode.ThreeWay)
            yield return "충돌";

        yield return "Mine (Dest/Curr)";
        if (comparisonMode == ComparisonMode.ThreeWay)
            yield return "Theirs (Src/Others)";
    }

    public string StartDiff3(string[] baseTextByLines, string[] mineTextByLines, string[] thierTextByLines)
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
