using XlsxMerge.Diff;
using XlsxMerge.Model;
using XlsxMerge.View;
using XlsxMerge.ViewModel;

namespace XlsxMerge
{
    internal static class Program
    {
        [STAThread]
        static int Main()
        {
            ApplicationConfiguration.Initialize();

            var args = Environment.GetCommandLineArgs();

            ProgramOptions? argumentInfo = null;
            if (args.Length > 1)
            {
                argumentInfo = ProgramOptions.Parse(args);
                if (argumentInfo == null || argumentInfo?.ComparisonMode == ComparisonMode.Unknown)
                {
                    argumentInfo = null;
                    MessageBox.Show("명령줄 인수에 잘못되거나 누락된 값이 있습니다.");
                }
            }

            // 폴더 변경은 args 해석 이후에 합니다.
            string? exeFolderPath = Path.GetDirectoryName(path: System.Reflection.Assembly.GetEntryAssembly()?.Location);
            if (String.IsNullOrEmpty(exeFolderPath) == false)
                Directory.SetCurrentDirectory(exeFolderPath);

            PathViewModel pathViewModel = new PathViewModel();

            if (argumentInfo != null)
            {
                pathViewModel.DiffPathModel = DiffPathModel.From(argumentInfo);
                var diffViewModel = new DiffViewModel();
                var formMainDiff = new FormMainDiff(pathViewModel, diffViewModel);
                Application.Run(formMainDiff);
                if (formMainDiff.MergeSuccessful)
                    return 0;
            }
            else
            {
                Application.Run(new FormWelcome(pathViewModel));
            }

            return 1;
        }
    }
}