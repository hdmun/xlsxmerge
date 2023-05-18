﻿using XlsxMerge.Diff;
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

            MergeArgumentInfo? argumentInfo = null;
            if (args.Length > 1)
            {
                argumentInfo = MergeArgumentInfo.Parse(args);
                if (argumentInfo.ComparisonMode == ComparisonMode.Unknown)
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
                var formMainDiff = new FormMainDiff(argumentInfo);
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