using System;
using System.IO;
using System.Linq;
using Nedev.FileConverters.DocxToDoc;

namespace Nedev.FileConverters.DocxToDoc.Cli
{
    internal class Program
    {
        private static int Main(string[] args)
        {
            var options = ParseArguments(args);

            if (options.ShowHelp || (string.IsNullOrEmpty(options.Input) && !options.BatchMode))
            {
                ShowHelp();
                return options.ShowHelp ? 0 : 1;
            }

            // Create logger based on verbose option
            ILogger logger = options.Verbose ? new ConsoleLogger() : NullLogger.Instance;

            try
            {
                var converter = new DocxToDocConverter(logger);

                if (options.BatchMode)
                {
                    return RunBatchConversion(options, converter, logger);
                }
                else
                {
                    return RunSingleConversion(options, converter, logger);
                }
            }
            catch (ConversionException ex)
            {
                logger.LogError($"Conversion failed at stage: {ex.Stage}", ex);
                Console.Error.WriteLine($"Conversion failed: {ex.Message}");
                if (options.Verbose && ex.InnerException != null)
                {
                    Console.Error.WriteLine($"Details: {ex.InnerException.Message}");
                }
                return 3;
            }
            catch (Exception ex)
            {
                logger.LogError("Unexpected error", ex);
                Console.Error.WriteLine("Conversion failed: " + ex.Message);
                if (options.Verbose)
                {
                    Console.Error.WriteLine(ex.StackTrace);
                }
                return 3;
            }
        }

        private static int RunSingleConversion(CliOptions options, DocxToDocConverter converter, ILogger logger)
        {
            if (!File.Exists(options.Input!))
            {
                logger.LogError($"Input file does not exist: '{options.Input}'");
                Console.Error.WriteLine($"Error: input file '{options.Input}' does not exist.");
                return 2;
            }

            string output = options.Output!;
            if (string.IsNullOrEmpty(output))
            {
                // Auto-generate output filename
                output = Path.ChangeExtension(options.Input!, ".doc");
            }

            logger.LogInfo($"Converting: {options.Input} -> {output}");

            if (options.Verbose)
            {
                Console.WriteLine($"Converting: {options.Input}");
                Console.WriteLine($"Output: {output}");
            }

            converter.Convert(options.Input!, output);

            if (options.Verbose)
            {
                var inputInfo = new FileInfo(options.Input!);
                var outputInfo = new FileInfo(output);
                Console.WriteLine($"Input size: {inputInfo.Length:N0} bytes");
                Console.WriteLine($"Output size: {outputInfo.Length:N0} bytes");
                Console.WriteLine($"Compression ratio: {(double)outputInfo.Length / inputInfo.Length:P1}");
            }

            logger.LogInfo($"Conversion completed: '{output}'");
            Console.WriteLine($"Converted '{options.Input}' -> '{output}'");
            return 0;
        }

        private static int RunBatchConversion(CliOptions options, DocxToDocConverter converter, ILogger logger)
        {
            if (!Directory.Exists(options.Input))
            {
                logger.LogError($"Input directory does not exist: '{options.Input}'");
                Console.Error.WriteLine($"Error: input directory '{options.Input}' does not exist.");
                return 2;
            }

            string searchPattern = options.Recursive ? "*.docx" : "*.docx";
            var searchOption = options.Recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

            var files = Directory.GetFiles(options.Input, searchPattern, searchOption)
                .Where(f => !options.ExcludeHidden || !new FileInfo(f).Attributes.HasFlag(FileAttributes.Hidden))
                .ToArray();

            if (files.Length == 0)
            {
                logger.LogInfo("No .docx files found in input directory");
                Console.WriteLine("No .docx files found.");
                return 0;
            }

            logger.LogInfo($"Found {files.Length} file(s) to convert in batch mode");

            Console.WriteLine($"Found {files.Length} file(s) to convert.");

            int successCount = 0;
            int failCount = 0;

            foreach (var file in files)
            {
                try
                {
                    string relativePath = Path.GetRelativePath(options.Input, file);
                    string outputFile;

                    if (!string.IsNullOrEmpty(options.Output))
                    {
                        // Preserve directory structure in output
                        string relativeDir = Path.GetDirectoryName(relativePath) ?? "";
                        string outputDir = Path.Combine(options.Output, relativeDir);
                        Directory.CreateDirectory(outputDir);
                        outputFile = Path.Combine(outputDir, Path.ChangeExtension(Path.GetFileName(file), ".doc"));
                    }
                    else
                    {
                        outputFile = Path.ChangeExtension(file, ".doc");
                    }

                    logger.LogInfo($"[{successCount + failCount + 1}/{files.Length}] Converting: {relativePath}");

                    if (options.Verbose)
                    {
                        Console.WriteLine($"[{successCount + failCount + 1}/{files.Length}] {relativePath}");
                    }

                    converter.Convert(file, outputFile);
                    successCount++;

                    logger.LogInfo($"  -> {outputFile}");

                    if (options.Verbose)
                    {
                        Console.WriteLine($"  -> {outputFile}");
                    }
                }
                catch (Exception ex)
                {
                    failCount++;
                    logger.LogError($"Failed to convert '{file}'", ex);
                    Console.Error.WriteLine($"Failed to convert '{file}': {ex.Message}");
                }
            }

            logger.LogInfo($"Batch conversion complete: {successCount} succeeded, {failCount} failed");
            Console.WriteLine($"\nConversion complete: {successCount} succeeded, {failCount} failed.");
            return failCount > 0 ? 4 : 0;
        }

        private static CliOptions ParseArguments(string[] args)
        {
            var options = new CliOptions();

            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i];

                switch (arg.ToLowerInvariant())
                {
                    case "-h":
                    case "--help":
                        options.ShowHelp = true;
                        break;

                    case "-v":
                    case "--verbose":
                        options.Verbose = true;
                        break;

                    case "-b":
                    case "--batch":
                        options.BatchMode = true;
                        break;

                    case "-r":
                    case "--recursive":
                        options.Recursive = true;
                        break;

                    case "--no-hidden":
                        options.ExcludeHidden = true;
                        break;

                    case "-o":
                    case "--output":
                        if (i + 1 < args.Length)
                        {
                            options.Output = args[++i];
                        }
                        break;

                    default:
                        if (!arg.StartsWith("-"))
                        {
                            if (options.Input == null)
                            {
                                options.Input = arg;
                            }
                            else if (options.Output == null)
                            {
                                options.Output = arg;
                            }
                        }
                        break;
                }
            }

            return options;
        }

        private static void ShowHelp()
        {
            Console.WriteLine("Nedev.FileConverters.DocxToDoc.Cli");
            Console.WriteLine("Convert OpenXML .docx files to legacy binary .doc format.");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  Single file:    dotnet cli.dll <input.docx> [output.doc]");
            Console.WriteLine("  Batch mode:     dotnet cli.dll -b <input-dir> [-o <output-dir>]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  -h, --help       Show this help message");
            Console.WriteLine("  -v, --verbose    Enable verbose output");
            Console.WriteLine("  -b, --batch      Enable batch mode (convert directory)");
            Console.WriteLine("  -r, --recursive  Process subdirectories recursively");
            Console.WriteLine("  --no-hidden      Exclude hidden files");
            Console.WriteLine("  -o, --output     Specify output file or directory");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  dotnet cli.dll document.docx");
            Console.WriteLine("  dotnet cli.dll document.docx output.doc");
            Console.WriteLine("  dotnet cli.dll -b ./documents -o ./converted -r -v");
            Console.WriteLine();
            Console.WriteLine("Exit codes:");
            Console.WriteLine("  0  Success");
            Console.WriteLine("  1  Invalid arguments");
            Console.WriteLine("  2  Input not found");
            Console.WriteLine("  3  Conversion error");
            Console.WriteLine("  4  Partial batch failure");
        }
    }

    internal class CliOptions
    {
        public string? Input { get; set; }
        public string? Output { get; set; }
        public bool ShowHelp { get; set; }
        public bool Verbose { get; set; }
        public bool BatchMode { get; set; }
        public bool Recursive { get; set; }
        public bool ExcludeHidden { get; set; }
    }
}
