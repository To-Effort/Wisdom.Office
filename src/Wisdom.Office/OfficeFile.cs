using System;
using System.IO;
using System.Text;
using Wisdom.Excel;
using Wisdom.Markdown;
using Wisdom.Office.Core;
using Wisdom.Pdf;
using Wisdom.Ppt;
using Wisdom.Word;

namespace Wisdom.Office
{
    /// <summary>
    /// Word、Excel、PowerPoint、PDF 与 Markdown 文档的统一访问入口。
    /// 通过 <c>Load</c> 系列静态方法打开任意受支持的文件，随后可提取文本、跨格式转换或另存为其他格式。
    /// Unified access entry point for Word, Excel, PowerPoint, PDF and Markdown documents.
    /// Use the <c>Load</c> static methods to open any supported file, then extract text, convert across formats, or save as another format.
    /// </summary>
    public sealed class OfficeFile
    {
        /// <summary>文档模型；打开 Excel 时惰性生成（首次调用 <see cref="AsDocument()"/> 时转换）。Document model; created lazily for Excel (converted on the first <see cref="AsDocument()"/> call).</summary>
        private OfficeDocument _document;
        /// <summary>Excel 工作簿内容；文件不是电子表格时为 null。Excel workbook content; null when the file is not a spreadsheet.</summary>
        private readonly OfficeWorkbook _workbook;
        /// <summary>PDF 原生文档对象；文件不是 PDF 时为 null。Native PDF document object; null when the file is not a PDF.</summary>
        private readonly PdfDocument _pdf;
        /// <summary>演示文稿内容；Word/PDF/Excel 载入后首次调用 <see cref="AsPresentation()"/> 时惰性生成。Presentation content; created lazily for Word/PDF/Excel (converted on the first <see cref="AsPresentation()"/> call).</summary>
        private OfficePresentation _presentation;

        /// <summary>私有构造函数，仅由本类的静态工厂方法调用，用于绑定格式与各类内容载体。Private constructor, called only by this class's static factory methods; binds the format and each content carrier.</summary>
        /// <param name="format">文件格式。File format.</param>
        /// <param name="document">文档模型或 null。Document model or null.</param>
        /// <param name="workbook">工作簿或 null。Workbook or null.</param>
        /// <param name="pdf">PDF 文档或 null。PDF document or null.</param>
        /// <param name="presentation">演示文稿或 null。Presentation or null.</param>
        private OfficeFile(OfficeFileFormat format, OfficeDocument document, OfficeWorkbook workbook, PdfDocument pdf, OfficePresentation presentation)
        {
            Format = format;
            _document = document;
            _workbook = workbook;
            _pdf = pdf;
            _presentation = presentation;
        }

        /// <summary>获取打开文件时识别出的 Office 格式。Gets the Office format detected when the file was opened.</summary>
        public OfficeFileFormat Format { get; private set; }

        /// <summary>获取文档模型；打开的是尚未转换的 Excel 文件时可能为 null，此时请改用 <see cref="AsDocument()"/>。Gets the document model; may be null for an unconverted Excel file, in which case use <see cref="AsDocument()"/> instead.</summary>
        public OfficeDocument Document
        {
            get { return _document; }
        }

        /// <summary>获取 Excel 工作簿；文件不是电子表格时返回 null。Gets the Excel workbook; returns null when the file is not a spreadsheet.</summary>
        public OfficeWorkbook Workbook
        {
            get { return _workbook; }
        }

        /// <summary>获取 PDF 原生文档；文件不是 PDF 时返回 null。Gets the native PDF document; returns null when the file is not a PDF.</summary>
        public PdfDocument Pdf
        {
            get { return _pdf; }
        }

        /// <summary>获取演示文稿；文件不是幻灯片且尚未转换时返回 null，此时请改用 <see cref="AsPresentation()"/>。Gets the presentation; may be null before conversion, in which case use <see cref="AsPresentation()"/> instead.</summary>
        public OfficePresentation Presentation
        {
            get { return _presentation; }
        }

        /// <summary>获取文档元数据（标题、作者、创建/修改时间等）；无法提取时返回空元数据对象。Gets document metadata (title, author, created/modified times, etc.); returns an empty metadata object when it cannot be extracted.</summary>
        public DocumentMetadata Metadata
        {
            get
            {
                if (_workbook != null) return _workbook.Metadata;
                if (_presentation != null) return _presentation.Metadata;
                if (_document != null) return _document.Metadata;
                return new DocumentMetadata();
            }
        }

        /// <summary>
        /// 从文件路径加载文档：先按文件内容（文件头）自动识别格式，识别失败时按扩展名兜底。
        /// 支持 PDF、DOCX、DOC、XLSX、XLS、PPTX、Markdown。
        /// Loads a document from a file path: auto-detects the format from the file content (header) first, falling back to the extension when detection fails.
        /// Supports PDF, DOCX, DOC, XLSX, XLS, PPTX and Markdown.
        /// </summary>
        /// <param name="path">文件路径。File path.</param>
        /// <returns>封装了文档内容的 <see cref="OfficeFile"/> 实例。An <see cref="OfficeFile"/> instance wrapping the document content.</returns>
        /// <remarks>示例：<c>OfficeFile doc = OfficeFile.Load(@"C:\data\report.xlsx");</c>，
        /// 随后可用 <see cref="ExtractText()"/> 提取文本，或用 <see cref="Save(string)"/> 另存为 PDF。
        /// Example: load a file, then extract text with <see cref="ExtractText()"/> or save as PDF with <see cref="Save(string)"/>.</remarks>
        public static OfficeFile Load(string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            byte[] data = File.ReadAllBytes(path);
            OfficeFileFormat format = OfficeFileFormats.Detect(data);
            if (format == OfficeFileFormat.Unknown)
            {
                format = OfficeFileFormats.FromExtension(Path.GetExtension(path));
            }
            return Load(data, format);
        }

        /// <summary>从流加载文档，先将流内容完整读入内存，再仅依据文件头识别格式。Loads a document from a stream; reads the entire stream into memory first, then detects the format from the file header only.</summary>
        /// <param name="stream">包含文档内容的可读流。A readable stream containing the document content.</param>
        /// <returns>封装了文档内容的 <see cref="OfficeFile"/> 实例。An <see cref="OfficeFile"/> instance wrapping the document content.</returns>
        public static OfficeFile Load(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            byte[] data = OfficeFileFormats.ReadAll(stream);
            return Load(data, OfficeFileFormats.Detect(data));
        }

        /// <summary>从字节数组加载文档，仅依据文件头识别格式。Loads a document from a byte array, detecting the format from the file header only.</summary>
        /// <param name="data">文档的字节内容。The byte content of the document.</param>
        /// <returns>封装了文档内容的 <see cref="OfficeFile"/> 实例。An <see cref="OfficeFile"/> instance wrapping the document content.</returns>
        public static OfficeFile Load(byte[] data)
        {
            if (data == null) throw new ArgumentNullException("data");
            return Load(data, OfficeFileFormats.Detect(data));
        }

        /// <summary>按指定格式从字节数组加载文档；格式为 <see cref="OfficeFileFormat.Unknown"/> 时自动识别。Loads a document from a byte array using the given format; auto-detects when the format is <see cref="OfficeFileFormat.Unknown"/>.</summary>
        /// <param name="data">文档的字节内容。The byte content of the document.</param>
        /// <param name="format">期望的文件格式。The expected file format.</param>
        /// <returns>封装了文档内容的 <see cref="OfficeFile"/> 实例。An <see cref="OfficeFile"/> instance wrapping the document content.</returns>
        /// <exception cref="NotSupportedException">内容无法按给定或识别出的格式解析时抛出。Thrown when the content cannot be parsed with the given or detected format.</exception>
        public static OfficeFile Load(byte[] data, OfficeFileFormat format)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (format == OfficeFileFormat.Unknown) format = OfficeFileFormats.Detect(data);

            switch (format)
            {
                case OfficeFileFormat.Pdf:
                    {
                        PdfDocument pdf = PdfDocument.Load(data);
                        return new OfficeFile(format, pdf.ToOfficeDocument(), null, pdf, null);
                    }
                case OfficeFileFormat.Docx:
                    return new OfficeFile(format, DocxReader.Load(data), null, null, null);
                case OfficeFileFormat.Dotx:
                    return new OfficeFile(format, DocxReader.Load(data), null, null, null);
                case OfficeFileFormat.Doc:
                    return new OfficeFile(format, DocReader.Load(data), null, null, null);
                case OfficeFileFormat.Rtf:
                    return new OfficeFile(format, RtfReader.Load(data), null, null, null);
                case OfficeFileFormat.Html:
                    return new OfficeFile(format, HtmlReader.Load(data), null, null, null);
                case OfficeFileFormat.Odt:
                    return new OfficeFile(format, OdtReader.Load(data), null, null, null);
                case OfficeFileFormat.Text:
                    return new OfficeFile(format, OfficeDocument.FromText(DecodeUtf8(data)), null, null, null);
                case OfficeFileFormat.Xlsx:
                    return new OfficeFile(format, null, XlsxReader.Load(data), null, null);
                case OfficeFileFormat.Xls:
                    return new OfficeFile(format, null, XlsReader.Load(data), null, null);
                case OfficeFileFormat.Pptx:
                    return new OfficeFile(format, null, null, null, PptxReader.Load(data));
                case OfficeFileFormat.Potx:
                    return new OfficeFile(format, null, null, null, PptxReader.Load(data));
                case OfficeFileFormat.Markdown:
                    return new OfficeFile(format, Wisdom.Markdown.Markdown.ToDocument(DecodeUtf8(data)), null, null, null);
                default:
                    throw new NotSupportedException("无法识别文件格式。支持 PDF、DOCX、DOC、XLSX、XLS、PPTX 和 Markdown。");
            }
        }

        /// <summary>以内存中的文档模型创建统一入口对象，初始格式视为 DOCX。Creates a unified entry object from an in-memory document model; the initial format is treated as DOCX.</summary>
        /// <param name="document">已构建好的 <see cref="OfficeDocument"/>。A prepared <see cref="OfficeDocument"/>.</param>
        /// <returns>封装该文档的 <see cref="OfficeFile"/> 实例。An <see cref="OfficeFile"/> instance wrapping the document.</returns>
        public static OfficeFile FromDocument(OfficeDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");
            return new OfficeFile(OfficeFileFormat.Docx, document, null, null, null);
        }

        /// <summary>以内存中的工作簿创建统一入口对象，初始格式视为 XLSX。Creates a unified entry object from an in-memory workbook; the initial format is treated as XLSX.</summary>
        /// <param name="workbook">已构建好的 <see cref="OfficeWorkbook"/>。A prepared <see cref="OfficeWorkbook"/>.</param>
        /// <returns>封装该工作簿的 <see cref="OfficeFile"/> 实例。An <see cref="OfficeFile"/> instance wrapping the workbook.</returns>
        public static OfficeFile FromWorkbook(OfficeWorkbook workbook)
        {
            if (workbook == null) throw new ArgumentNullException("workbook");
            return new OfficeFile(OfficeFileFormat.Xlsx, null, workbook, null, null);
        }

        /// <summary>以内存中的演示文稿创建统一入口对象，初始格式视为 PPTX。Creates a unified entry object from an in-memory presentation; the initial format is treated as PPTX.</summary>
        /// <param name="presentation">已构建好的 <see cref="OfficePresentation"/>。A prepared <see cref="OfficePresentation"/>.</param>
        /// <returns>封装该演示文稿的 <see cref="OfficeFile"/> 实例。An <see cref="OfficeFile"/> instance wrapping the presentation.</returns>
        public static OfficeFile FromPresentation(OfficePresentation presentation)
        {
            if (presentation == null) throw new ArgumentNullException("presentation");
            return new OfficeFile(OfficeFileFormat.Pptx, null, null, null, presentation);
        }

        /// <summary>获取文档模型：Word/PDF 直接返回；Excel 按默认排版规则转换为文档并缓存结果。Gets the document model: returned directly for Word/PDF; for Excel converted to a document using default layout rules and the result cached.</summary>
        /// <returns>文档模型；文件既无文档也无工作簿内容时返回 null。The document model; null when the file has neither document nor workbook content.</returns>
        public OfficeDocument AsDocument()
        {
            return AsDocument(null);
        }

        /// <summary>获取文档模型；Excel 工作簿按指定排版选项转换（传入 options 时不缓存转换结果）。Gets the document model; Excel workbooks are converted using the given layout options (the result is not cached when options are supplied).</summary>
        /// <param name="options">工作表转文档的排版选项，null 表示使用默认设置。Layout options for converting worksheets to a document; null uses default settings.</param>
        /// <returns>文档模型；文件不含文档/工作簿内容时返回 null。The document model; null when the file contains no document/workbook content.</returns>
        public OfficeDocument AsDocument(WorkbookLayoutOptions options)
        {
            if (_document != null) return _document;
            if (_workbook != null)
            {
                OfficeDocument converted = WorkbookToDocumentConverter.Convert(_workbook, options);
                if (options == null) _document = converted;
                return converted;
            }
            if (_presentation != null)
            {
                _document = PresentationToDocumentConverter.Convert(_presentation);
                return _document;
            }
            return null;
        }

        /// <summary>获取演示文稿：PPTX 直接返回；Word/Excel/PDF 文档按标题层级拆分幻灯片并缓存转换结果。Gets the presentation: returned directly for PPTX; Word/Excel/PDF documents are split into slides by heading level and the result is cached.</summary>
        /// <returns>演示文稿；文件既无演示文稿也无可用文档内容时返回 null。The presentation; null when the file has neither a presentation nor usable document content.</returns>
        public OfficePresentation AsPresentation()
        {
            if (_presentation != null) return _presentation;
            OfficeDocument document = AsDocument();
            if (document == null) return null;
            _presentation = DocumentToPresentationConverter.Convert(document);
            return _presentation;
        }

        /// <summary>提取文档纯文本：PDF 读取文字层，Word 取段落文本，Excel 按"工作表名 + 制表符分隔的行列"展开，PPTX 按"第 N 页 + 形状文本"展开。Extracts plain text: reads the PDF text layer, the Word paragraph text, expands Excel as "sheet name + tab-separated rows and columns", and expands PPTX as "slide N + shape text".</summary>
        /// <returns>纯文本内容；无可用内容时返回空字符串。Plain text content; empty string when no content is available.</returns>
        public string ExtractText()
        {
            if (_pdf != null) return _pdf.ExtractText();
            if (_document != null) return _document.GetPlainText();
            if (_workbook != null) return ExtractWorkbookText(_workbook);
            if (_presentation != null) return ExtractPresentationText(_presentation);
            return string.Empty;
        }

        /// <summary>保存到指定路径，输出格式按目标扩展名决定；扩展名无法识别时沿用当前格式。Saves to the given path; the output format is decided by the target extension; falls back to the current format when the extension is unrecognized.</summary>
        /// <param name="path">目标文件路径。The target file path.</param>
        public void Save(string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            OfficeFileFormat target = OfficeFileFormats.FromExtension(Path.GetExtension(path));
            if (target == OfficeFileFormat.Unknown) target = Format;
            Save(path, target);
        }

        /// <summary>以指定格式将内容保存到文件路径。Saves the content to a file path in the given format.</summary>
        /// <param name="path">目标文件路径。The target file path.</param>
        /// <param name="format">输出格式。The output format.</param>
        public void Save(string path, OfficeFileFormat format)
        {
            if (path == null) throw new ArgumentNullException("path");
            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                Save(stream, format);
            }
        }

        /// <summary>以指定格式将内容写入流；这是跨格式转换的核心方法（如 XLSX 转 PDF、DOCX 转 XLSX 场景中的输出环节）。Writes the content to a stream in the given format; the core method for cross-format conversion (e.g. the output step of XLSX-to-PDF or DOCX-to-XLSX).</summary>
        /// <param name="stream">可写目标流。A writable target stream.</param>
        /// <param name="format">输出格式，支持 PDF、DOCX、XLSX、PPTX、Markdown。The output format; PDF, DOCX, XLSX, PPTX and Markdown are supported.</param>
        /// <remarks>旧版 .doc/.xls 只支持读取，写入时会抛出 <see cref="NotSupportedException"/>；
        /// 电子表格以外的内容（Word、PDF、PPTX、Markdown）不能直接存为 XLSX，其余组合同样会抛 <see cref="NotSupportedException"/>。
        /// Legacy .doc/.xls are read-only and throw <see cref="NotSupportedException"/> on write; non-spreadsheet content
        /// (Word, PDF, PPTX, Markdown) cannot be saved directly as XLSX, and any other unsupported pairing also throws.</remarks>
        public void Save(Stream stream, OfficeFileFormat format)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            switch (format)
            {
                case OfficeFileFormat.Pdf:
                    // 幻灯片内容走固定版式渲染，保留形状位置与样式；其余格式仍按流式文档模型排版。
                    // Slide content uses fixed-layout rendering to preserve shape position and style; other
                    // sources still lay out through the flow document model.
                    if (_presentation != null)
                    {
                        new PdfRenderer().RenderPresentation(_presentation, stream);
                        break;
                    }
                    new PdfRenderer().Render(RequireDocument(), stream);
                    break;
                case OfficeFileFormat.Docx:
                    DocxWriter.Write(RequireDocument(), stream);
                    break;
                case OfficeFileFormat.Dotx:
                    DocxWriter.WriteTemplate(RequireDocument(), stream);
                    break;
                case OfficeFileFormat.Rtf:
                    RtfWriter.Write(RequireDocument(), stream);
                    break;
                case OfficeFileFormat.Html:
                    HtmlWriter.Write(RequireDocument(), stream);
                    break;
                case OfficeFileFormat.Odt:
                    OdtWriter.Write(RequireDocument(), stream);
                    break;
                case OfficeFileFormat.Text:
                    {
                        byte[] text = new UTF8Encoding(false).GetBytes(RequireDocument().GetPlainText());
                        stream.Write(text, 0, text.Length);
                        break;
                    }
                case OfficeFileFormat.Xlsx:
                    if (_workbook == null)
                    {
                        throw new NotSupportedException("Word/PDF 文档不能直接保存为 Excel 工作簿。");
                    }
                    XlsxWriter.Write(_workbook, stream);
                    break;
                case OfficeFileFormat.Pptx:
                    PptxWriter.Write(RequirePresentation(), stream);
                    break;
                case OfficeFileFormat.Potx:
                    PptxWriter.WriteTemplate(RequirePresentation(), stream);
                    break;
                case OfficeFileFormat.Markdown:
                    {
                        byte[] text = new UTF8Encoding(false).GetBytes(Wisdom.Markdown.Markdown.ToMarkdown(RequireDocument()));
                        stream.Write(text, 0, text.Length);
                        break;
                    }
                case OfficeFileFormat.Doc:
                case OfficeFileFormat.Xls:
                    throw new NotSupportedException("旧版 .doc/.xls 格式只支持读取，请使用 .docx/.xlsx 保存。");
                default:
                    throw new NotSupportedException("不支持的输出格式：" + format);
            }
        }

        /// <summary>将文件转换/序列化为指定格式的字节数组，适合 Web 下载等无文件场景。Converts/serializes the file to a byte array in the given format, suitable for file-less scenarios such as web downloads.</summary>
        /// <param name="format">输出格式，支持 PDF、DOCX、XLSX、PPTX、Markdown。The output format; PDF, DOCX, XLSX, PPTX and Markdown are supported.</param>
        /// <returns>目标格式的文件字节。The file bytes in the target format.</returns>
        public byte[] ToArray(OfficeFileFormat format)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                Save(stream, format);
                return stream.ToArray();
            }
        }

        /// <summary>内部辅助：确保存在可用的文档模型（Excel 会先转换），不存在时抛出 <see cref="NotSupportedException"/>。Internal helper: ensures a usable document model exists (Excel is converted first); throws <see cref="NotSupportedException"/> when none exists.</summary>
        /// <returns>非 null 的文档模型。A non-null document model.</returns>
        private OfficeDocument RequireDocument()
        {
            OfficeDocument document = AsDocument();
            if (document == null)
            {
                throw new NotSupportedException("当前文件没有可用的文档内容。");
            }
            return document;
        }

        /// <summary>内部辅助：确保存在可用的演示文稿（文档会先按标题拆分），PDF 内容不支持转换时抛出 <see cref="NotSupportedException"/>。Internal helper: ensures a usable presentation exists (documents are split by heading first); throws <see cref="NotSupportedException"/> for content that cannot be converted.</summary>
        /// <returns>非 null 的演示文稿。A non-null presentation.</returns>
        private OfficePresentation RequirePresentation()
        {
            OfficePresentation presentation = AsPresentation();
            if (presentation == null)
            {
                throw new NotSupportedException("当前文件没有可用的演示文稿内容。");
            }
            return presentation;
        }

        /// <summary>内部辅助：按 UTF-8 解码字节并去掉起始 BOM。Internal helper: decodes bytes as UTF-8 and removes a leading BOM.</summary>
        /// <param name="data">字节内容。The byte content.</param>
        /// <returns>解码后的文本。The decoded text.</returns>
        private static string DecodeUtf8(byte[] data)
        {
            int offset = 0;
            if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF) offset = 3;
            return new UTF8Encoding(false).GetString(data, offset, data.Length - offset);
        }

        /// <summary>将工作簿逐表、逐行展开为以制表符分隔的纯文本（跳过隐藏工作表）。Expands the workbook sheet by sheet, row by row into tab-separated plain text (skipping hidden worksheets).</summary>
        /// <param name="workbook">要提取文本的工作簿。The workbook to extract text from.</param>
        /// <returns>展开后的纯文本。The expanded plain text.</returns>
        private static string ExtractWorkbookText(OfficeWorkbook workbook)
        {
            StringBuilder builder = new StringBuilder();
            for (int s = 0; s < workbook.Worksheets.Count; s++)
            {
                Worksheet sheet = workbook.Worksheets[s];
                if (sheet.IsHidden) continue;
                if (builder.Length > 0) builder.AppendLine();
                builder.Append('[').Append(sheet.Name).Append(']').AppendLine();
                XlsxReader.AppendSheetText(builder, sheet);
            }
            return builder.ToString();
        }

        /// <summary>将演示文稿逐页展开为纯文本，每页以"第 N 页：标题"开头，随后按形状输出文本或图片/表格占位说明。Expands the presentation page by page into plain text; each page starts with "Slide N: title" followed by each shape's text or an image/table placeholder note.</summary>
        /// <param name="presentation">要提取文本的演示文稿。The presentation to extract text from.</param>
        /// <returns>展开后的纯文本。The expanded plain text.</returns>
        private static string ExtractPresentationText(OfficePresentation presentation)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                Slide slide = presentation.Slides[i];
                if (builder.Length > 0) builder.AppendLine();
                string title = slide.GetTitle();
                builder.Append("第 ").Append(i + 1).Append(" 页");
                if (!string.IsNullOrEmpty(title)) builder.Append('：').Append(title);
                builder.AppendLine();
                for (int s = 0; s < slide.Shapes.Count; s++)
                {
                    SlideShape shape = slide.Shapes[s];
                    TextBoxShape text = shape as TextBoxShape;
                    if (text != null)
                    {
                        string value = text.GetText().Trim();
                        if (value.Length > 0 && value != title) builder.AppendLine(value);
                        continue;
                    }
                    PictureShape picture = shape as PictureShape;
                    if (picture != null)
                    {
                        builder.AppendLine("[图片] " + picture.AlternativeText);
                        continue;
                    }
                    TableShape table = shape as TableShape;
                    if (table != null && table.Table != null)
                    {
                        builder.Append(ExtractTableText(table.Table));
                    }
                }
            }
            return builder.ToString();
        }

        /// <summary>将表格按行输出，同行单元格以制表符分隔。Outputs a table row by row, cells within a row separated by tabs.</summary>
        /// <param name="table">要提取文本的表格。The table to extract text from.</param>
        /// <returns>表格纯文本。The table's plain text.</returns>
        private static string ExtractTableText(Table table)
        {
            StringBuilder builder = new StringBuilder();
            for (int r = 0; r < table.Rows.Count; r++)
            {
                TableRow row = table.Rows[r];
                for (int c = 0; c < row.Cells.Count; c++)
                {
                    if (c > 0) builder.Append('\t');
                    for (int b = 0; b < row.Cells[c].Blocks.Count; b++)
                    {
                        Paragraph paragraph = row.Cells[c].Blocks[b] as Paragraph;
                        if (paragraph != null) builder.Append(paragraph.GetText());
                    }
                }
                if (builder.Length > 0) builder.AppendLine();
            }
            return builder.ToString();
        }
    }
}
