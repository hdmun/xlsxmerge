using System.Diagnostics;
using System.Text;

namespace XlsxMerge.Features.Diffs;

public static class Diff3Process
{
    private static readonly string FileName = @".\diff3.exe";

    public static string CreateTempFile(string[]? texts)
    {
        string temp = Path.GetTempFileName();
        if (texts != null)
            File.WriteAllLines(temp, texts);
        return temp;
    }

    public static string Start(string[] diffFiles)
    {
        ProcessStartInfo psi = new ProcessStartInfo()
        {
            FileName = Path.GetFullPath(FileName),
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            StandardOutputEncoding = Encoding.UTF8,
            Arguments = diffFiles.Aggregate("", (prev, current) => $"{prev} \"{current}\"").Trim(),
        };
        psi.WorkingDirectory = Path.GetDirectoryName(psi.FileName);

        var p = Process.Start(psi);
        string result = p.StandardOutput.ReadToEnd();
        p.WaitForExit();

        foreach (var path in diffFiles)
        {
            File.Delete(path);
        }

        return result;
    }
}
