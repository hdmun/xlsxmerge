using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;
using XlsxMerge.Extensions;
using XlsxMerge.Model;

namespace XlsxMerge.ViewModel;

public class PathViewModel : INotifyPropertyChanged
{
    public DiffPathModel DiffPathModel { get; private set; }

    public PathViewModel()
    {
        DiffPathModel = new DiffPathModel();
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public List<string>? Arguments(bool use3wayMerge)
    {
        var args = new List<string>
        {
            $"{Assembly.GetEntryAssembly()?.GetName().Name}.exe",
            $"-b {BasePath.AddDoubleQuote()}",
            $"-d {MinePath.AddDoubleQuote()}"
        };
        if (use3wayMerge)
            args.Add($"-s {TheirsPath.AddDoubleQuote()}");

        if (!string.IsNullOrEmpty(ResultPath))
            args.Add($"-r {ResultPath.AddDoubleQuote()}");

        string resultArgs = string.Join(" ", args);
        if (resultArgs.Contains("\"\""))
            return null;

        return args;
    }

    public string BasePath
    {
        get => DiffPathModel.BasePath;
        set
        {
            DiffPathModel.BasePath = value;
            OnPropertyChanged();
        }
    }

    public string MinePath
    {
        get => DiffPathModel.MinePath;
        set
        {
            DiffPathModel.MinePath = value;
            OnPropertyChanged();
        }
    }

    public string TheirsPath
    {
        get => DiffPathModel.TheirsPath;
        set
        {
            DiffPathModel.TheirsPath = value;
            OnPropertyChanged();
        }
    }

    public string ResultPath
    {
        get => DiffPathModel.ResultPath;
        set
        {
            DiffPathModel.ResultPath = value;
            OnPropertyChanged();
        }
    }

    private void OnPropertyChanged([CallerMemberName] string propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
