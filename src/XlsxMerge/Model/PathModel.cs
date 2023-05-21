namespace XlsxMerge.Model;

public class DiffPathModel
{
    public static DiffPathModel From(ProgramOptions args)
    {
        return new DiffPathModel
        {
            BasePath = args.BasePath,
            MinePath = args.MinePath,
            TheirsPath = args.TheirsPath,
            ResultPath = args.ResultPath,
        };
    }

    public string BasePath { get; set; } = string.Empty;
    public string MinePath { get; set; } = string.Empty;
    public string TheirsPath { get; set; } = string.Empty;
    public string ResultPath { get; set; } = string.Empty;
}
