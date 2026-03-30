using System;
using System.IO;
using System.Diagnostics;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests
{
    public class CliTests
    {
        private string GetCliPath()
        {
            var possiblePaths = new[]
            {
                Path.Combine(AppContext.BaseDirectory, "Nedev.FileConverters.DocxToDoc.Cli.dll"),
                Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Nedev.FileConverters.DocxToDoc.Cli", "bin", "Debug", "net8.0", "Nedev.FileConverters.DocxToDoc.Cli.dll"),
                Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Nedev.FileConverters.DocxToDoc.Cli", "bin", "Release", "net8.0", "Nedev.FileConverters.DocxToDoc.Cli.dll")
            };

            foreach (var path in possiblePaths)
            {
                var fullPath = Path.GetFullPath(path);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }

            throw new FileNotFoundException("CLI assembly was not found. Ensure test project references and builds the CLI project.");
        }

        private ProcessStartInfo CreateStartInfo(string arguments)
        {
            var cliPath = GetCliPath();
            return new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = $"\"{cliPath}\" {arguments}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
        }

        [Fact]
        public void Cli_Help_ReturnsZero()
        {
            var psi = CreateStartInfo("--help");

            using var process = Process.Start(psi);
            Assert.NotNull(process);
            process!.WaitForExit();
            string output = process.StandardOutput.ReadToEnd();

            Assert.Equal(0, process.ExitCode);
            Assert.Contains("Usage:", output);
            Assert.Contains("Options:", output);
        }

        [Fact]
        public void Cli_NoArguments_ReturnsError()
        {
            var psi = CreateStartInfo("");

            using var process = Process.Start(psi);
            Assert.NotNull(process);
            process!.WaitForExit();

            Assert.Equal(1, process.ExitCode);
        }

        [Fact]
        public void Cli_NonExistentFile_ReturnsError()
        {
            var psi = CreateStartInfo("nonexistent.docx output.doc");

            using var process = Process.Start(psi);
            Assert.NotNull(process);
            process!.WaitForExit();
            string error = process.StandardError.ReadToEnd();

            Assert.Equal(2, process.ExitCode);
            Assert.Contains("does not exist", error);
        }

        [Fact]
        public void Cli_VerboseFlag_ShowsDetails()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);
            string inputFile = Path.Combine(tempDir, "test.docx");
            string outputFile = Path.Combine(tempDir, "test.doc");

            try
            {
                CreateMinimalDocx(inputFile);

                var psi = CreateStartInfo($"-v \"{inputFile}\" \"{outputFile}\"");
                using var process = Process.Start(psi);
                Assert.NotNull(process);
                process!.WaitForExit();
                string output = process.StandardOutput.ReadToEnd();

                Assert.Equal(0, process.ExitCode);
                Assert.Contains("Converting:", output);
                Assert.Contains("bytes", output);
            }
            finally
            {
                // Cleanup
                try { Directory.Delete(tempDir, true); } catch { }
            }
        }

        private void CreateMinimalDocx(string path)
        {
            using var fs = File.Create(path);
            using var archive = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Create);

            var entry = archive.CreateEntry("word/document.xml");
            using var stream = entry.Open();
            using var writer = new StreamWriter(stream);
            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                "<w:body><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:body></w:document>");
        }
    }
}
