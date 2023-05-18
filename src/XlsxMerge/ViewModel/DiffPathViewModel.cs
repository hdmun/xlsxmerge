using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;
using XlsxMerge.Extensions;

namespace XlsxMerge.ViewModel;

public class DiffPathViewModel : INotifyPropertyChanged
{
    private string _basePath;
    private string _minePath;
    private string _theirsPath;
    private string _resultPath;

    public DiffPathViewModel()
    {
        _basePath = string.Empty;
        _minePath = string.Empty;
        _theirsPath = string.Empty;
        _resultPath = string.Empty;
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public List<string>? Arguments(bool use3wayMerge)
    {
        var args = new List<string>
        {
            $"{Assembly.GetEntryAssembly()?.GetName().Name}.exe",
            $"-b={BasePath.AddDoubleQuote()}",
            $"-d={MinePath.AddDoubleQuote()}"
        };
        if (use3wayMerge)
            args.Add($"-s={TheirsPath.AddDoubleQuote()}");

        if (!string.IsNullOrEmpty(ResultPath))
            args.Add($"-r={ResultPath.AddDoubleQuote()}");

        string resultArgs = string.Join(" ", args);
        if (resultArgs.Contains("=\"\""))
            return null;

        return args;
    }

    public string BasePath
    {
        get => _basePath;
        set
        {
            _basePath = value;
            OnPropertyChanged();
        }
    }

    public string MinePath
    {
        get => _minePath;
        set
        {
            _minePath = value;
            OnPropertyChanged();
        }
    }

    public string TheirsPath
    {
        get => _theirsPath;
        set
        {
            _theirsPath = value;
            OnPropertyChanged();
        }
    }

    public string ResultPath
    {
        get => _resultPath;
        set
        {
            _resultPath = value;
            OnPropertyChanged();
        }
    }

    private void OnPropertyChanged([CallerMemberName] string propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
