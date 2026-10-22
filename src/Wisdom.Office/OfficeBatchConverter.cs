using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Wisdom.Office.Core;

namespace Wisdom.Office
{
    /// <summary>一项批量转换任务：源路径、目标路径与可选的工作表排版选项。One batch-conversion job: source path, target path and optional worksheet layout options.</summary>
    public sealed class BatchConversionItem
    {
        /// <summary>源文件路径。Source file path.</summary>
        public string SourcePath { get; set; }
        /// <summary>目标文件路径，其扩展名决定输出格式。Target file path; its extension decides the output format.</summary>
        public string TargetPath { get; set; }
        /// <summary>工作表排版选项，可为 null。Worksheet layout options; may be null.</summary>
        public WorkbookLayoutOptions Options { get; set; }

        /// <summary>创建空任务。Creates an empty job.</summary>
        public BatchConversionItem() { }

        /// <summary>创建一项转换任务。Creates one conversion job.</summary>
        /// <param name="sourcePath">源文件路径。Source file path.</param>
        /// <param name="targetPath">目标文件路径。Target file path.</param>
        /// <param name="options">工作表排版选项，可为 null。Worksheet layout options; may be null.</param>
        public BatchConversionItem(string sourcePath, string targetPath, WorkbookLayoutOptions options)
        {
            SourcePath = sourcePath;
            TargetPath = targetPath;
            Options = options;
        }
    }

    /// <summary>一项批量任务的执行结果。 The outcome of a single batch job.</summary>
    public sealed class BatchConversionOutcome
    {
        /// <summary>对应的任务项。 The originating job.</summary>
        public BatchConversionItem Item { get; set; }
        /// <summary>是否成功。 Whether the job succeeded.</summary>
        public bool Success { get; set; }
        /// <summary>失败时捕获的异常；成功为 null。The captured exception on failure; null on success.</summary>
        public Exception Error { get; set; }
    }

    /// <summary>整批转换的汇总报告。 The summary report for a whole batch.</summary>
    public sealed class BatchConversionReport
    {
        /// <summary>任务总数。 Total number of jobs.</summary>
        public int Total { get; set; }
        /// <summary>成功数。 Number of successes.</summary>
        public int Succeeded { get; set; }
        /// <summary>失败数。 Number of failures.</summary>
        public int Failed { get; set; }
        /// <summary>耗时（毫秒）。Elapsed time in milliseconds.</summary>
        public long ElapsedMilliseconds { get; set; }
        /// <summary>逐项结果，按输入顺序排列。 Per-item outcomes, in input order.</summary>
        public IReadOnlyList<BatchConversionOutcome> Outcomes { get; set; }

        /// <summary>创建汇总报告。Creates a summary report.</summary>
        public BatchConversionReport()
        {
            Outcomes = new List<BatchConversionOutcome>();
        }
    }

    /// <summary>
    /// 企业级批量与异步转换门面：把 <see cref="OfficeConverter"/> 的单文件同步转换组合成
    /// 有界并行的整批转换、目录扫描转换与基于 <see cref="Task"/> 的异步封装。各模块已确认线程安全，
    /// 因此并行度只受 <c>maxDegreeOfParallelism</c> 与取消令牌约束，单个文件失败不影响整批。
    /// Enterprise batch and async façade: composes the single-file synchronous conversions of
    /// <see cref="OfficeConverter"/> into bounded-parallel batch runs, directory scans and <see cref="Task"/>-based
    /// async wrappers. The modules are thread-safe, so parallelism is bounded only by
    /// <c>maxDegreeOfParallelism</c> and the cancellation token, and one failing file never aborts the batch.
    /// </summary>
    public static class OfficeBatchConverter
    {
        /// <summary>把默认并行度限制为不超过该上限，避免超大请求把线程池打满。 Caps the default parallelism so oversized requests do not swamp the thread pool.</summary>
        private const int DefaultParallelismCap = 16;

        /// <summary>有界并行地执行整批转换；返回逐项结果汇总，不抛出单项异常。 Runs the whole batch with bounded parallelism; returns a per-item report and does not rethrow individual failures.</summary>
        /// <param name="items">任务项集合。 The jobs.</param>
        /// <param name="maxDegreeOfParallelism">最大并行度，非正数时按 CPU 核数（上限 16）。Maximum parallelism; when not positive uses the core count capped at 16.</param>
        /// <param name="progress">逐项完成回调（可选），在任意工作线程上调用。 Optional per-item completion callback invoked on a worker thread.</param>
        /// <param name="cancellationToken">取消令牌。 Cancellation token.</param>
        /// <returns>整批汇总报告。 The batch report.</returns>
        /// <exception cref="OperationCanceledException">取消令牌触发时抛出。Thrown when the cancellation token fires.</exception>
        public static BatchConversionReport Convert(
            IEnumerable<BatchConversionItem> items,
            int maxDegreeOfParallelism = 0,
            IProgress<BatchConversionOutcome> progress = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (items == null) throw new ArgumentNullException("items");
            List<BatchConversionItem> list = items.ToList();
            BatchConversionOutcome[] slots = new BatchConversionOutcome[list.Count];
            int dop = ResolveDegree(maxDegreeOfParallelism);
            long started = Environment.TickCount;

            ParallelOptions options = new ParallelOptions();
            options.MaxDegreeOfParallelism = dop;
            options.CancellationToken = cancellationToken;
            Parallel.ForEach(Partitioned(list), options, delegate(int[] pair)
            {
                cancellationToken.ThrowIfCancellationRequested();
                BatchConversionOutcome outcome = RunItem(list[pair[0]], pair[0]);
                slots[pair[0]] = outcome;
                if (progress != null) progress.Report(outcome);
            });

            return BuildReport(slots, Environment.TickCount - started);
        }

        /// <summary>在线程池上异步执行整批有界并行转换。 Runs the whole batch asynchronously on the thread pool with bounded parallelism.</summary>
        /// <param name="items">任务项集合。 The jobs.</param>
        /// <param name="maxDegreeOfParallelism">最大并行度，非正数时按 CPU 核数（上限 16）。Maximum parallelism; when not positive uses the core count capped at 16.</param>
        /// <param name="progress">逐项完成回调（可选）。 Optional per-item completion callback.</param>
        /// <param name="cancellationToken">取消令牌。 Cancellation token.</param>
        /// <returns>表示整批完成的 <see cref="Task{BatchConversionReport}"/>。 A task yielding the batch report.</returns>
        public static Task<BatchConversionReport> ConvertAsync(
            IEnumerable<BatchConversionItem> items,
            int maxDegreeOfParallelism = 0,
            IProgress<BatchConversionOutcome> progress = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            // 快照到列表，避免在工作线程上延迟枚举非线程安全的调用方序列。
            // Materialize so a caller's non-thread-safe sequence is not enumerated on worker threads.
            List<BatchConversionItem> snapshot = items == null ? null : items.ToList();
            if (snapshot == null) throw new ArgumentNullException("items");
            return Task.Run(() => Convert(snapshot, maxDegreeOfParallelism, progress, cancellationToken), cancellationToken);
        }

        /// <summary>异步转换单个文件（在 <see cref="Task"/> 上运行同步转换），便于 await 而非阻塞调用线程。 Converts one file asynchronously on a <see cref="Task"/> so callers await instead of blocking.</summary>
        /// <param name="sourcePath">源文件路径。 Source file path.</param>
        /// <param name="targetPath">目标文件路径。 Target file path.</param>
        /// <param name="options">工作表排版选项，可为 null。 Worksheet layout options; may be null.</param>
        /// <param name="cancellationToken">取消令牌。 Cancellation token.</param>
        /// <returns>表示完成 <see cref="Task"/>。 A task representing completion.</returns>
        public static Task ConvertFileAsync(
            string sourcePath,
            string targetPath,
            WorkbookLayoutOptions options = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            return Task.Run(delegate
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (options == null) OfficeConverter.Convert(sourcePath, targetPath);
                else OfficeConverter.Convert(sourcePath, targetPath, options);
            }, cancellationToken);
        }

        /// <summary>异步把已加载文件另存为目标格式。 Saves an already-loaded file in the target format asynchronously.</summary>
        /// <param name="file">已加载文件。 The loaded file.</param>
        /// <param name="targetPath">目标文件路径。 Target file path.</param>
        /// <param name="options">工作表排版选项，可为 null。 Worksheet layout options; may be null.</param>
        /// <param name="cancellationToken">取消令牌。 Cancellation token.</param>
        /// <returns>表示完成 <see cref="Task"/>。 A task representing completion.</returns>
        public static Task ConvertAsync(
            OfficeFile file,
            string targetPath,
            WorkbookLayoutOptions options = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            return Task.Run(delegate
            {
                cancellationToken.ThrowIfCancellationRequested();
                OfficeConverter.Convert(file, targetPath, options);
            }, cancellationToken);
        }

        /// <summary>异步把内存中的文档模型渲染为 PDF 字节数组。 Renders an in-memory document model to PDF bytes asynchronously.</summary>
        /// <param name="document">文档模型。 The document model.</param>
        /// <param name="cancellationToken">取消令牌。 Cancellation token.</param>
        /// <returns>PDF 字节数组的 <see cref="Task{Byte}"/>。 A task yielding the PDF bytes.</returns>
        public static Task<byte[]> ToPdfAsync(OfficeDocument document, CancellationToken cancellationToken = default(CancellationToken))
        {
            return Task.Run(delegate
            {
                cancellationToken.ThrowIfCancellationRequested();
                return OfficeConverter.ToPdf(document);
            }, cancellationToken);
        }

        /// <summary>扫描源目录并以有界并行把每个匹配文件转换到目标目录的相同相对路径。 Scans a source directory and converts every matching file into the same relative path under a target directory, with bounded parallelism.</summary>
        /// <param name="sourceDirectory">源目录。 Source directory.</param>
        /// <param name="targetDirectory">目标目录（不存在则创建）。 Target directory, created when missing.</param>
        /// <param name="targetExtension">目标扩展名，如 <c>.pdf</c>。 Target extension, such as <c>.pdf</c>.</param>
        /// <param name="maxDegreeOfParallelism">最大并行度，非正数时按 CPU 核数（上限 16）。 Maximum parallelism; when not positive uses the core count capped at 16.</param>
        /// <param name="searchOption">是否递归子目录。 Whether to recurse into subdirectories.</param>
        /// <param name="cancellationToken">取消令牌。 Cancellation token.</param>
        /// <returns>整批汇总报告。 The batch report.</returns>
        public static BatchConversionReport ConvertDirectory(
            string sourceDirectory,
            string targetDirectory,
            string targetExtension,
            int maxDegreeOfParallelism = 0,
            SearchOption searchOption = SearchOption.TopDirectoryOnly,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            return Convert(BuildDirectoryItems(sourceDirectory, targetDirectory, targetExtension, searchOption),
                maxDegreeOfParallelism, null, cancellationToken);
        }

        /// <summary>异步版本的目录扫描并行转换。 Async version of the directory-scan parallel conversion.</summary>
        /// <param name="sourceDirectory">源目录。 Source directory.</param>
        /// <param name="targetDirectory">目标目录。 Target directory.</param>
        /// <param name="targetExtension">目标扩展名。 Target extension.</param>
        /// <param name="maxDegreeOfParallelism">最大并行度。 Maximum parallelism.</param>
        /// <param name="searchOption">是否递归子目录。 Whether to recurse.</param>
        /// <param name="cancellationToken">取消令牌。 Cancellation token.</param>
        /// <returns>整批汇总报告的 <see cref="Task{BatchConversionReport}"/>。 A task yielding the batch report.</returns>
        public static Task<BatchConversionReport> ConvertDirectoryAsync(
            string sourceDirectory,
            string targetDirectory,
            string targetExtension,
            int maxDegreeOfParallelism = 0,
            SearchOption searchOption = SearchOption.TopDirectoryOnly,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            List<BatchConversionItem> items = BuildDirectoryItems(sourceDirectory, targetDirectory, targetExtension, searchOption);
            return ConvertAsync(items, maxDegreeOfParallelism, null, cancellationToken);
        }

        /// <summary>把源目录下所有受支持文件映射为以目标扩展名重命名的目标任务项。 Maps every supported file under the source directory into target jobs renamed with the target extension.</summary>
        /// <param name="sourceDirectory">源目录。 Source directory.</param>
        /// <param name="targetDirectory">目标目录。 Target directory.</param>
        /// <param name="targetExtension">目标扩展名。 Target extension.</param>
        /// <param name="searchOption">是否递归子目录。 Whether to recurse.</param>
        /// <returns>任务项列表。 The jobs.</returns>
        private static List<BatchConversionItem> BuildDirectoryItems(
            string sourceDirectory, string targetDirectory, string targetExtension, SearchOption searchOption)
        {
            if (sourceDirectory == null) throw new ArgumentNullException("sourceDirectory");
            if (targetDirectory == null) throw new ArgumentNullException("targetDirectory");
            if (string.IsNullOrEmpty(targetExtension)) throw new ArgumentException("targetExtension is required", "targetExtension");
            if (!Directory.Exists(sourceDirectory)) throw new DirectoryNotFoundException(sourceDirectory);
            if (!targetExtension.StartsWith(".", StringComparison.Ordinal)) targetExtension = "." + targetExtension;

            string sourceRoot = Path.GetFullPath(sourceDirectory);
            string targetRoot = Path.GetFullPath(targetDirectory);
            Directory.CreateDirectory(targetRoot);

            string[] files = Directory.GetFiles(sourceRoot, "*.*", searchOption);
            List<BatchConversionItem> items = new List<BatchConversionItem>();
            for (int i = 0; i < files.Length; i++)
            {
                string sourcePath = files[i];
                OfficeFileFormat format = OfficeFileFormats.FromExtension(Path.GetExtension(sourcePath));
                if (format == OfficeFileFormat.Unknown) continue;
                string relative = sourcePath.Substring(sourceRoot.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                string changed = Path.ChangeExtension(relative, null) + targetExtension;
                string targetPath = Path.Combine(targetRoot, changed);
                string parent = Path.GetDirectoryName(targetPath);
                if (!string.IsNullOrEmpty(parent)) Directory.CreateDirectory(parent);
                items.Add(new BatchConversionItem(sourcePath, targetPath, null));
            }
            return items;
        }

        /// <summary>执行单个任务项，捕获异常而不抛出。 Runs a single job, capturing any exception instead of throwing.</summary>
        /// <param name="item">任务项。 The job.</param>
        /// <param name="unusedIndex">占位序号（保留供调试）。 Placeholder index, kept for debugging.</param>
        /// <returns>逐项结果。 The outcome.</returns>
        private static BatchConversionOutcome RunItem(BatchConversionItem item, int unusedIndex)
        {
            BatchConversionOutcome outcome = new BatchConversionOutcome();
            outcome.Item = item;
            try
            {
                OfficeConverter.Convert(item.SourcePath, item.TargetPath, item.Options);
                outcome.Success = true;
            }
            catch (Exception error)
            {
                outcome.Success = false;
                outcome.Error = error;
            }
            return outcome;
        }

        /// <summary>把逐项结果数组折叠为汇总报告。 Folds the outcome array into a summary report.</summary>
        /// <param name="slots">按输入顺序填充的结果数组。 The results, filled in input order.</param>
        /// <param name="elapsedMs">耗时毫秒。 Elapsed milliseconds.</param>
        /// <returns>汇总报告。 The report.</returns>
        private static BatchConversionReport BuildReport(BatchConversionOutcome[] slots, long elapsedMs)
        {
            BatchConversionReport report = new BatchConversionReport();
            List<BatchConversionOutcome> list = new List<BatchConversionOutcome>(slots.Length);
            int succeeded = 0;
            for (int i = 0; i < slots.Length; i++)
            {
                BatchConversionOutcome outcome = slots[i];
                if (outcome == null)
                {
                    outcome = new BatchConversionOutcome();
                    outcome.Success = false;
                    outcome.Error = new OperationCanceledException("任务未完成即取消。 The job was cancelled before it ran.");
                }
                if (outcome.Success) succeeded++;
                list.Add(outcome);
            }
            report.Total = slots.Length;
            report.Succeeded = succeeded;
            report.Failed = slots.Length - succeeded;
            report.Outcomes = list;
            report.ElapsedMilliseconds = elapsedMs < 0 ? 0 : elapsedMs;
            return report;
        }

        /// <summary>解析请求的并行度：非正数用 CPU 核数，统一封顶。 Resolve the requested parallelism: a non-positive value uses the core count, always capped.</summary>
        /// <param name="requested">请求的并行度。 The requested degree.</param>
        /// <returns>实际并行度。 The effective degree.</returns>
        private static int ResolveDegree(int requested)
        {
            int cores = Environment.ProcessorCount;
            if (cores < 1) cores = 1;
            if (cores > DefaultParallelismCap) cores = DefaultParallelismCap;
            if (requested <= 0) return cores;
            return requested > DefaultParallelismCap * 8 ? DefaultParallelismCap * 8 : requested;
        }

        /// <summary>把列表切成 (索引, 索引) 单元素数组以便 Parallel.ForEach 定位结果槽位。 Slices the list into single-element index arrays so Parallel.ForEach can place each outcome into its slot.</summary>
        /// <param name="list">输入列表。 The input list.</param>
        /// <returns>可并行的索引序列。 A parallelizable sequence of indices.</returns>
        private static IEnumerable<int[]> Partitioned(List<BatchConversionItem> list)
        {
            for (int i = 0; i < list.Count; i++)
            {
                yield return new int[] { i };
            }
        }
    }
}
