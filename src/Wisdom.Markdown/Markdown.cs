using System;
using System.IO;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Markdown
{
    /// <summary>
    /// Markdown 门面：把 Markdown 文本解析为统一文档模型 <see cref="OfficeDocument"/>，
    /// 并在文档模型与 HTML/Markdown 之间互转。
    /// 解析结果可直接交给 Wisdom.Word / Wisdom.Pdf 等导出器生成 Word 或 PDF。
    /// The Markdown facade: parses Markdown text into the unified document model <see cref="OfficeDocument"/>
    /// and converts between the document model and HTML/Markdown.
    /// The parsed result can be handed directly to exporters such as Wisdom.Word or Wisdom.Pdf to produce Word or PDF.
    /// </summary>
    public static class Markdown
    {
        /// <summary>
        /// 使用默认选项把 Markdown 文本解析为文档模型。
        /// Parses Markdown text into the document model using default options.
        /// </summary>
        /// <param name="markdown">
        /// Markdown 源文本，null 或空文本返回仅含空节的合法文档。The Markdown source text; null or empty input returns a valid document with only an empty section.
        /// </param>
        /// <returns>
        /// 解析得到的 <see cref="OfficeDocument"/>。The parsed <see cref="OfficeDocument"/>.
        /// </returns>
        public static OfficeDocument ToDocument(string markdown)
        {
            return ToDocument(markdown, null);
        }

        /// <summary>
        /// 使用指定选项把 Markdown 文本解析为文档模型。
        /// Parses Markdown text into the document model using the given options.
        /// </summary>
        /// <param name="markdown">
        /// Markdown 源文本。The Markdown source text.
        /// </param>
        /// <param name="options">
        /// 解析选项，null 时使用默认选项。The parse options; default options are used when null.
        /// </param>
        /// <returns>
        /// 解析得到的 <see cref="OfficeDocument"/>。The parsed <see cref="OfficeDocument"/>.
        /// </returns>
        public static OfficeDocument ToDocument(string markdown, MarkdownOptions options)
        {
            MarkdownParser parser = new MarkdownParser(options);
            return parser.Parse(markdown);
        }

        /// <summary>
        /// 按 UTF-8 读取 Markdown 文件并自动去除 BOM，解析为文档模型（使用默认选项）。
        /// Reads a Markdown file as UTF-8, automatically strips the BOM, and parses it into the document model using default options.
        /// </summary>
        /// <param name="path">
        /// Markdown 文件路径。The Markdown file path.
        /// </param>
        /// <returns>
        /// 解析得到的 <see cref="OfficeDocument"/>。The parsed <see cref="OfficeDocument"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException">
        /// <paramref name="path"/> 为 null。Is null.
        /// </exception>
        public static OfficeDocument LoadFile(string path)
        {
            return LoadFile(path, null);
        }

        /// <summary>
        /// 按 UTF-8 读取 Markdown 文件并自动去除 BOM，使用指定选项解析为文档模型。
        /// Reads a Markdown file as UTF-8, automatically strips the BOM, and parses it into the document model using the given options.
        /// </summary>
        /// <param name="path">
        /// Markdown 文件路径。The Markdown file path.
        /// </param>
        /// <param name="options">
        /// 解析选项，null 时使用默认选项。The parse options; default options are used when null.
        /// </param>
        /// <returns>
        /// 解析得到的 <see cref="OfficeDocument"/>。The parsed <see cref="OfficeDocument"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException">
        /// <paramref name="path"/> 为 null。Is null.
        /// </exception>
        public static OfficeDocument LoadFile(string path, MarkdownOptions options)
        {
            if (path == null)
            {
                throw new ArgumentNullException("path");
            }
            string text = ReadAllTextUtf8(path);
            return ToDocument(text, WithFileBaseDirectory(path, options));
        }

        /// <summary>
        /// 让从文件加载的 Markdown 默认把图片解析限定在该文件所在目录内；调用方显式指定过基准目录时保持不变，且不会修改传入的选项实例。
        /// Lets Markdown loaded from a file resolve local images inside that file's own directory by default; an explicitly
        /// set base directory is kept, and the supplied options instance is never mutated.
        /// </summary>
        /// <param name="path">Markdown 文件路径。The Markdown file path.</param>
        /// <param name="options">调用方选项，可能为 null。The caller's options, possibly null.</param>
        /// <returns>生效选项。The effective options.</returns>
        private static MarkdownOptions WithFileBaseDirectory(string path, MarkdownOptions options)
        {
            MarkdownOptions effective = options == null ? new MarkdownOptions() : options.Clone();
            if (string.IsNullOrEmpty(effective.ImageBaseDirectory))
            {
                try
                {
                    string directory = Path.GetDirectoryName(Path.GetFullPath(path));
                    if (!string.IsNullOrEmpty(directory))
                    {
                        effective.ImageBaseDirectory = directory;
                    }
                }
                catch (Exception)
                {
                    // 路径无法归一化时保持禁用本地图片。When the path cannot be normalized, local images stay disabled.
                }
            }
            return effective;
        }

        /// <summary>
        /// 使用默认选项把 Markdown 文本渲染为 HTML 片段。
        /// Renders Markdown text into an HTML fragment using default options.
        /// </summary>
        /// <param name="markdown">
        /// Markdown 源文本。The Markdown source text.
        /// </param>
        /// <returns>
        /// 渲染得到的 HTML 字符串。The rendered HTML string.
        /// </returns>
        public static string ToHtml(string markdown)
        {
            return ToHtml(markdown, null);
        }

        /// <summary>
        /// 使用指定选项把 Markdown 文本渲染为 HTML 片段（先解析为文档模型，再交由 HTML 写器渲染）。
        /// Renders Markdown text into an HTML fragment using the given options (parsing into the document model first, then rendering via the HTML writer).
        /// </summary>
        /// <param name="markdown">
        /// Markdown 源文本。The Markdown source text.
        /// </param>
        /// <param name="options">
        /// 渲染选项，null 时使用默认选项。The render options; default options are used when null.
        /// </param>
        /// <returns>
        /// 渲染得到的 HTML 字符串。The rendered HTML string.
        /// </returns>
        public static string ToHtml(string markdown, MarkdownOptions options)
        {
            MarkdownOptions effective = options ?? new MarkdownOptions();
            OfficeDocument document = ToDocument(markdown, effective);
            return MarkdownHtmlWriter.Render(document, effective);
        }

        /// <summary>
        /// 使用默认选项把文档模型回写为 Markdown 文本。
        /// Writes the document model back to Markdown text using default options.
        /// </summary>
        /// <param name="document">
        /// 待回写的文档，null 返回空字符串。The document to write back; null returns an empty string.
        /// </param>
        /// <returns>
        /// 生成的 Markdown 文本。The generated Markdown text.
        /// </returns>
        public static string ToMarkdown(OfficeDocument document)
        {
            return ToMarkdown(document, null);
        }

        /// <summary>
        /// 使用指定选项把文档模型回写为 Markdown 文本（交由 Markdown 写器渲染）。
        /// Writes the document model back to Markdown text using the given options (rendering via the Markdown writer).
        /// </summary>
        /// <param name="document">
        /// 待回写的文档，null 返回空字符串。The document to write back; null returns an empty string.
        /// </param>
        /// <param name="options">
        /// 渲染选项，null 时使用默认选项。The render options; default options are used when null.
        /// </param>
        /// <returns>
        /// 生成的 Markdown 文本。The generated Markdown text.
        /// </returns>
        public static string ToMarkdown(OfficeDocument document, MarkdownOptions options)
        {
            if (document == null)
            {
                return string.Empty;
            }
            MarkdownOptions effective = options ?? new MarkdownOptions();
            return OfficeDocumentMarkdownWriter.Render(document, effective);
        }

        /// <summary>单个 Markdown 源文件允许读取的最大字节数（64 MB）。The largest Markdown source file that may be read (64 MB).</summary>
        private const int MaxMarkdownBytes = 64 * 1024 * 1024;

        /// <summary>
        /// 以 UTF-8 读取全部文本并去除起始 BOM；超过大小上限的文件抛 <see cref="InvalidDataException"/>，避免超大输入耗尽内存。
        /// Reads all text as UTF-8 and removes a leading BOM; a file beyond the size cap throws an <see cref="InvalidDataException"/>
        /// instead of exhausting memory.
        /// </summary>
        /// <param name="path">
        /// 文件路径。The file path.
        /// </param>
        /// <returns>
        /// 去除 BOM 后的文本内容。The text content without the BOM.
        /// </returns>
        private static string ReadAllTextUtf8(string path)
        {
            long length = new FileInfo(path).Length;
            if (length > MaxMarkdownBytes)
            {
                throw new InvalidDataException("Markdown 文件超过 " + MaxMarkdownBytes + " 字节上限。/ The Markdown file exceeds the " + MaxMarkdownBytes + "-byte limit.");
            }
            byte[] bytes = File.ReadAllBytes(path);
            Encoding utf8 = new UTF8Encoding(false);
            int offset = 0;
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            {
                offset = 3;
            }
            return utf8.GetString(bytes, offset, bytes.Length - offset);
        }
    }
}
