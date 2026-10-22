using System;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>Word 文件级整理结果：输入/输出/备份/报告路径与整理数据。The file-level clean-up result: input, output, backup and report paths plus the normalization data.</summary>
    public sealed class WordProcessResult
    {
        /// <summary>输入文档路径。The path of the input document.</summary>
        public string InputPath { get; set; }
        /// <summary>规范化输出 .docx 路径。The path of the normalized .docx output.</summary>
        public string OutputPath { get; set; }
        /// <summary>原文件备份路径。The backup path of the original file.</summary>
        public string BackupPath { get; set; }
        /// <summary>整理报告文本文件路径。The path of the written report file.</summary>
        public string ReportPath { get; set; }
        /// <summary>整理报告全文文本。The full report text.</summary>
        public string Report { get; set; }
        /// <summary>规范化管道结果。The normalization pipeline result.</summary>
        public WordNormalizeResult Normalize { get; set; }
        /// <summary>整理前后差异对比。The before/after document diff.</summary>
        public WordDocumentDiff Diff { get; set; }
    }

    /// <summary>
    /// 文件级入口：加载 .doc/.docx，规范化文档并写出规范 .docx，保留原文件备份并可选写出整理报告。
    /// File-level entry point: loads .doc/.docx, normalizes the document, writes a normalized .docx,
    /// keeps a backup of the original and optionally writes a clean-up report.
    /// </summary>
    public static class WordProcessor
    {
        /// <summary>
        /// 执行文件级整理：加载文档、规范化、写出规范 .docx，并生成备份与报告文件。
        /// Runs the file-level clean-up: loads the document, normalizes it, writes the normalized .docx
        /// and produces the backup and report files.
        /// </summary>
        /// <param name="inputPath">输入 .doc/.docx 路径。The input .doc/.docx path.</param>
        /// <param name="outputPath">输出 .docx 路径。The output .docx path.</param>
        /// <param name="options">规范化选项，为 null 时使用默认。The normalization options; defaults are used when null.</param>
        /// <returns>文件整理结果。The file-level result.</returns>
        public static WordProcessResult Process(string inputPath, string outputPath, WordNormalizeOptions options)
        {
            if (inputPath == null) throw new ArgumentNullException("inputPath");
            if (outputPath == null) throw new ArgumentNullException("outputPath");
            if (!File.Exists(inputPath)) throw new FileNotFoundException("Input document was not found.", inputPath);
            if (options == null) options = new WordNormalizeOptions();

            string backupPath = CreateBackup(inputPath);

            OfficeDocument document = LoadDocument(inputPath);
            OfficeDocument original = DocumentCloner.Clone(document);

            WordNormalizeResult normalizeResult = WordNormalizer.Normalize(document, options);
            WordDocumentDiff diff = WordDocumentDiff.Compare(original, document);
            normalizeResult.Diff = diff.ToText();

            string outputDirectory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDirectory) && !Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }
            DocxWriter.Write(document, outputPath);

            string report = normalizeResult.ToReport();
            string reportPath = outputPath + ".report.txt";
            File.WriteAllText(reportPath, report, new UTF8Encoding(false));

            WordProcessResult result = new WordProcessResult();
            result.InputPath = inputPath;
            result.OutputPath = outputPath;
            result.BackupPath = backupPath;
            result.ReportPath = reportPath;
            result.Report = report;
            result.Normalize = normalizeResult;
            result.Diff = diff;
            return result;
        }

        /// <summary>
        /// 在源文件旁生成带时间戳的备份副本，重名时自动追加序号。
        /// Creates a timestamped backup copy next to the source file, appending a counter when the name collides.
        /// </summary>
        /// <param name="path">待备份文件路径。The file to back up.</param>
        /// <returns>实际创建的备份路径。The backup path actually created.</returns>
        public static string CreateBackup(string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            string directory = Path.GetDirectoryName(path);
            string name = Path.GetFileNameWithoutExtension(path);
            string extension = Path.GetExtension(path);
            string stamp = DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
            string candidate = Path.Combine(directory ?? string.Empty, name + ".backup-" + stamp + extension);
            int suffix = 2;
            while (File.Exists(candidate))
            {
                candidate = Path.Combine(directory ?? string.Empty,
                    name + ".backup-" + stamp + "-" + suffix.ToString(CultureInfo.InvariantCulture) + extension);
                suffix++;
            }
            File.Copy(path, candidate);
            return candidate;
        }

        /// <summary>
        /// 用备份文件覆盖目标文件。
        /// Restores the target file by overwriting it with the backup.
        /// </summary>
        /// <param name="backupPath">备份文件路径。The backup file path.</param>
        /// <param name="targetPath">覆盖目标路径。The target to restore.</param>
        public static void Rollback(string backupPath, string targetPath)
        {
            if (backupPath == null) throw new ArgumentNullException("backupPath");
            if (targetPath == null) throw new ArgumentNullException("targetPath");
            if (!File.Exists(backupPath)) throw new FileNotFoundException("Backup file was not found.", backupPath);
            File.Copy(backupPath, targetPath, true);
        }

        private static OfficeDocument LoadDocument(string path)
        {
            string extension = (Path.GetExtension(path) ?? string.Empty).ToLowerInvariant();
            if (extension == ".docx" || extension == ".docm")
            {
                return DocxReader.Load(path);
            }
            if (extension == ".doc")
            {
                return DocReader.Load(path);
            }
            if (extension.Length == 0)
            {
                return DocxReader.Load(path);
            }
            throw new NotSupportedException("Unsupported document format: " + extension);
        }
    }
}