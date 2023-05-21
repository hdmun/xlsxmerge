using CommandLine;
using XlsxMerge.Diff;

namespace XlsxMerge
{
    public class ProgramOptions
    {
        public static ProgramOptions Parse(string[] args)
        {
            var result = Parser.Default.ParseArguments<ProgramOptions>(args);
            return result.Value;
        }

        [Option('b', "base", Required = true, HelpText = "" )]
        public string? BasePath { get; set; }

        [Option('d', "diff", Required = true, HelpText = "")]
        public string? MinePath { get; set; }

        [Option('s', "source", HelpText = " ")]
        public string? TheirsPath { get; set; }

        [Option('r', "result", Required = true, HelpText = "")]
        public string? ResultPath { get; set; }

        [Option('x', "extra", HelpText = "")]
        public string? ExtraInfoPath { get; set; }

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
    }
}
