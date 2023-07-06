using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;
using XlsxMerge.Extensions;
using XlsxMerge.Features.Diffs;
using XlsxMerge.Model;

namespace XlsxMerge.ViewModel;

public class PathViewModel : INotifyPropertyChanged
{
    public DiffPathModel DiffPathModel { get; set; }

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
            $"-b {BasePath}",
            $"-d {MinePath}"
        };
        if (use3wayMerge)
            args.Add($"-s {TheirsPath}");

        if (!string.IsNullOrEmpty(ResultPath))
            args.Add($"-r {ResultPath}");

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
            OnPropertyChanged(nameof(BasePathLabelText));
            OnPropertyChanged(nameof(VisibleTheirsPath));
            OnPropertyChanged(nameof(EnableDiffStart));
        }
    }

    public string MinePath
    {
        get => DiffPathModel.MinePath;
        set
        {
            DiffPathModel.MinePath = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(MinePathLabelText));
            OnPropertyChanged(nameof(VisibleTheirsPath));
            OnPropertyChanged(nameof(EnableDiffStart));
        }
    }

    public string TheirsPath
    {
        get => DiffPathModel.TheirsPath;
        set
        {
            DiffPathModel.TheirsPath = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(TheirsPathLabelText));
            OnPropertyChanged(nameof(VisibleTheirsPath));
        }
    }

    public string ResultPath
    {
        get => DiffPathModel.ResultPath;
        set
        {
            DiffPathModel.ResultPath = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(ResultPathLabelText));
            OnPropertyChanged(nameof(VisibleResultPath));
            OnPropertyChanged(nameof(EnableDiffStart));
        }
    }

    public string BasePathLabelText
    {
        get => $"Base : {Path.GetFileName(BasePath)} ({BasePath})";
    }

    public string MinePathLabelText
    {
        get => $"Mine (Destination, Current) : {Path.GetFileName(MinePath)} ({MinePath})";
    }

    public string TheirsPathLabelText
    {
        get => $"Theirs (Source, Others) : {Path.GetFileName(TheirsPath)} ({TheirsPath})";
    }

    public string ResultPathLabelText
    {
        get => $"Result : {Path.GetFileName(ResultPath)} ({ResultPath})";
    }

    public bool VisibleTheirsPath
    {
        get => ComparisonMode == ComparisonMode.ThreeWay;
    }

    public bool VisibleResultPath
    {
        get => string.IsNullOrEmpty(ResultPath) == false;
    }

    public bool EnableDiffStart
    {
        get => !string.IsNullOrEmpty(BasePath) && !string.IsNullOrEmpty(MinePath) && VisibleResultPath;
    }

    public ComparisonMode ComparisonMode
    {
        get
        {
            if (string.IsNullOrEmpty(BasePath) || string.IsNullOrEmpty(MinePath))
                return ComparisonMode.Unknown;

            if (string.IsNullOrEmpty(TheirsPath))
                return ComparisonMode.TwoWay;

            return ComparisonMode.ThreeWay;
        }
    }

    private void OnPropertyChanged([CallerMemberName] string propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
