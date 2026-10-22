using System;
using System.IO;
using System.Text;
using Wisdom.Markdown;
using Wisdom.Office.Core;
using Wisdom.Pdf;
using Wisdom.Ppt;
using Wisdom.Word;

namespace Wisdom.Office
{
    /// <summary>
    /// Office 格式互转工具：所有格式先载入统一文档模型（或演示文稿模型），再渲染为目标格式。
    /// 常用于"任意文件另存为另一种格式"或"把内存中的文档/工作簿/演示文稿直接输出为 PDF/Word/PPTX/Markdown"。
    /// Office format conversion utility: every format is loaded into the unified document model (or presentation model) first, then rendered to the target format.
    /// Commonly used to save any file as another format, or to output an in-memory document/workbook/presentation directly as PDF/Word/PPTX/Markdown.
    /// </summary>
    public static class OfficeConverter
    {
        /// <summary>将源文件转换为路径扩展名指定的目标格式文件。Converts the source file to a target-format file determined by the path extension.</summary>
        /// <param name="sourcePath">源文件路径（支持 PDF、DOCX、DOC、XLSX、XLS、PPTX、Markdown）。Source file path (PDF, DOCX, DOC, XLSX, XLS, PPTX and Markdown supported).</param>
        /// <param name="targetPath">目标文件路径，其扩展名决定输出格式。Target file path; its extension decides the output format.</param>
        /// <remarks>示例：<c>OfficeConverter.Convert("report.xlsx", "report.pdf");</c> 即可将 Excel 转成 PDF。
        /// Example: <c>OfficeConverter.Convert("report.xlsx", "report.pdf");</c> converts Excel to PDF.</remarks>
        public static void Convert(string sourcePath, string targetPath)
        {
            Convert(sourcePath, targetPath, null);
        }

        /// <summary>将源文件转换为目标格式文件；当源为 Excel 时可用排版选项控制版面。Converts the source file to a target-format file; layout options control the page layout when the source is Excel.</summary>
        /// <param name="sourcePath">源文件路径。Source file path.</param>
        /// <param name="targetPath">目标文件路径，其扩展名决定输出格式。Target file path; its extension decides the output format.</param>
        /// <param name="workbookLayout">工作表排版选项，null 表示使用默认设置。Worksheet layout options; null uses default settings.</param>
        public static void Convert(string sourcePath, string targetPath, WorkbookLayoutOptions workbookLayout)
        {
            if (sourcePath == null) throw new ArgumentNullException("sourcePath");
            if (targetPath == null) throw new ArgumentNullException("targetPath");
            OfficeFile file = OfficeFile.Load(sourcePath);
            Convert(file, targetPath, workbookLayout);
        }

        /// <summary>将已加载的文件另存为路径扩展名指定的格式。Saves the loaded file in the format indicated by the path extension.</summary>
        /// <param name="file">已加载的 <see cref="OfficeFile"/>。A loaded <see cref="OfficeFile"/>.</param>
        /// <param name="targetPath">目标文件路径。Target file path.</param>
        public static void Convert(OfficeFile file, string targetPath)
        {
            Convert(file, targetPath, null);
        }

        /// <summary>将已加载的文件另存为目标格式；输出 PDF/DOCX 且源为 Excel 时，用排版选项控制版面。Saves the loaded file in the target format; layout options control the page layout when outputting PDF/DOCX from an Excel source.</summary>
        /// <param name="file">已加载的 <see cref="OfficeFile"/>。A loaded <see cref="OfficeFile"/>.</param>
        /// <param name="targetPath">目标文件路径，其扩展名决定输出格式。Target file path; its extension decides the output format.</param>
        /// <param name="workbookLayout">工作表排版选项，null 表示使用默认设置。Worksheet layout options; null uses default settings.</param>
        /// <exception cref="NotSupportedException">目标路径扩展名无法识别为受支持的输出格式时抛出。Thrown when the target path extension is not a recognized supported output format.</exception>
        public static void Convert(OfficeFile file, string targetPath, WorkbookLayoutOptions workbookLayout)
        {
            if (file == null) throw new ArgumentNullException("file");
            if (targetPath == null) throw new ArgumentNullException("targetPath");

            OfficeFileFormat target = OfficeFileFormats.FromExtension(Path.GetExtension(targetPath));
            if (target == OfficeFileFormat.Unknown)
            {
                throw new NotSupportedException("无法根据扩展名确定输出格式：" + targetPath);
            }

            if (workbookLayout == null)
            {
                file.Save(targetPath, target);
                return;
            }

            switch (target)
            {
                case OfficeFileFormat.Pdf:
                    new PdfRenderer().Render(file.AsDocument(workbookLayout), targetPath);
                    break;
                case OfficeFileFormat.Docx:
                    DocxWriter.Write(file.AsDocument(workbookLayout), targetPath);
                    break;
                default:
                    file.Save(targetPath, target);
                    break;
            }
        }

        /// <summary>将源字节内容转换为目标格式的字节数组（格式自动识别）。Converts the source byte content to a byte array in the target format (format auto-detected).</summary>
        /// <param name="data">源文件字节。The source file bytes.</param>
        /// <param name="targetFormat">目标格式。The target format.</param>
        /// <returns>目标格式的文件字节。The file bytes in the target format.</returns>
        public static byte[] Convert(byte[] data, OfficeFileFormat targetFormat)
        {
            if (data == null) throw new ArgumentNullException("data");
            return OfficeFile.Load(data).ToArray(targetFormat);
        }

        /// <summary>将内存中的文档模型渲染为 PDF 文件。Renders the in-memory document model to a PDF file.</summary>
        /// <param name="document">文档模型。The document model.</param>
        /// <param name="path">PDF 输出路径。The PDF output path.</param>
        public static void ToPdf(OfficeDocument document, string path)
        {
            new PdfRenderer().Render(document, path);
        }

        /// <summary>将内存中的文档模型渲染为 PDF 字节数组。Renders the in-memory document model to a PDF byte array.</summary>
        /// <param name="document">文档模型。The document model.</param>
        /// <returns>PDF 文件字节。The PDF file bytes.</returns>
        public static byte[] ToPdf(OfficeDocument document)
        {
            return new PdfRenderer().RenderToBytes(document);
        }

        /// <summary>将内存中的文档模型保存为 DOCX（Word）文件。Saves the in-memory document model as a DOCX (Word) file.</summary>
        /// <param name="document">文档模型。The document model.</param>
        /// <param name="path">DOCX 输出路径。The DOCX output path.</param>
        public static void ToWord(OfficeDocument document, string path)
        {
            DocxWriter.Write(document, path);
        }

        /// <summary>将内存中的文档模型保存为 DOCX（Word）字节数组。Saves the in-memory document model as a DOCX (Word) byte array.</summary>
        /// <param name="document">文档模型。The document model.</param>
        /// <returns>DOCX 文件字节。The DOCX file bytes.</returns>
        public static byte[] ToWord(OfficeDocument document)
        {
            return DocxWriter.Write(document);
        }

        /// <summary>将工作簿按默认排版渲染为 PDF 文件（内部先转换为文档模型再渲染）。Renders the workbook to a PDF file using the default layout (internally converts to a document model first, then renders).</summary>
        /// <param name="workbook">要导出的 <see cref="OfficeWorkbook"/>。The <see cref="OfficeWorkbook"/> to export.</param>
        /// <param name="path">PDF 输出路径。The PDF output path.</param>
        public static void WorkbookToPdf(OfficeWorkbook workbook, string path)
        {
            WorkbookToPdf(workbook, path, null);
        }

        /// <summary>将工作簿按指定排版渲染为 PDF 文件。Renders the workbook to a PDF file using the given layout.</summary>
        /// <param name="workbook">要导出的 <see cref="OfficeWorkbook"/>。The <see cref="OfficeWorkbook"/> to export.</param>
        /// <param name="path">PDF 输出路径。The PDF output path.</param>
        /// <param name="options">工作表排版选项，null 表示使用默认设置。Worksheet layout options; null uses default settings.</param>
        public static void WorkbookToPdf(OfficeWorkbook workbook, string path, WorkbookLayoutOptions options)
        {
            new PdfRenderer().Render(WorkbookToDocumentConverter.Convert(workbook, options), path);
        }

        /// <summary>将工作簿按指定排版渲染为 PDF 字节数组。Renders the workbook to a PDF byte array using the given layout.</summary>
        /// <param name="workbook">要导出的 <see cref="OfficeWorkbook"/>。The <see cref="OfficeWorkbook"/> to export.</param>
        /// <param name="options">工作表排版选项，null 表示使用默认设置。Worksheet layout options; null uses default settings.</param>
        /// <returns>PDF 文件字节。The PDF file bytes.</returns>
        public static byte[] WorkbookToPdf(OfficeWorkbook workbook, WorkbookLayoutOptions options)
        {
            return new PdfRenderer().RenderToBytes(WorkbookToDocumentConverter.Convert(workbook, options));
        }

        /// <summary>将演示文稿渲染为 PDF 文件：固定版式逐页绘制，保留形状位置、填充、描边与文本排版。Renders the presentation to a PDF file in fixed layout, painting each slide as one page with shape positions, fills, outlines and text layout preserved.</summary>
        /// <param name="presentation">要导出的 <see cref="OfficePresentation"/>。The <see cref="OfficePresentation"/> to export.</param>
        /// <param name="path">PDF 输出路径。The PDF output path.</param>
        public static void PresentationToPdf(OfficePresentation presentation, string path)
        {
            if (presentation == null) throw new ArgumentNullException("presentation");
            new PdfRenderer().RenderPresentation(presentation, path);
        }

        /// <summary>将演示文稿按固定版式渲染为 PDF 字节数组。Renders the presentation to a PDF byte array in fixed layout.</summary>
        /// <param name="presentation">要导出的 <see cref="OfficePresentation"/>。The <see cref="OfficePresentation"/> to export.</param>
        /// <returns>PDF 文件字节。The PDF file bytes.</returns>
        public static byte[] PresentationToPdf(OfficePresentation presentation)
        {
            if (presentation == null) throw new ArgumentNullException("presentation");
            return new PdfRenderer().RenderPresentationToBytes(presentation);
        }

        /// <summary>将演示文稿保存为 DOCX（Word）文件，每页幻灯片对应文档中的一节。Saves the presentation as a DOCX (Word) file, one document section per slide.</summary>
        /// <param name="presentation">要导出的 <see cref="OfficePresentation"/>。The <see cref="OfficePresentation"/> to export.</param>
        /// <param name="path">DOCX 输出路径。The DOCX output path.</param>
        public static void PresentationToWord(OfficePresentation presentation, string path)
        {
            if (presentation == null) throw new ArgumentNullException("presentation");
            DocxWriter.Write(PresentationToDocumentConverter.Convert(presentation), path);
        }

        /// <summary>将演示文稿保存为 DOCX（Word）字节数组。Saves the presentation as a DOCX (Word) byte array.</summary>
        /// <param name="presentation">要导出的 <see cref="OfficePresentation"/>。The <see cref="OfficePresentation"/> to export.</param>
        /// <returns>DOCX 文件字节。The DOCX file bytes.</returns>
        public static byte[] PresentationToWord(OfficePresentation presentation)
        {
            if (presentation == null) throw new ArgumentNullException("presentation");
            return DocxWriter.Write(PresentationToDocumentConverter.Convert(presentation));
        }

        /// <summary>将演示文稿保存为 PPTX 文件。Saves the presentation as a PPTX file.</summary>
        /// <param name="presentation">要导出的 <see cref="OfficePresentation"/>。The <see cref="OfficePresentation"/> to export.</param>
        /// <param name="path">PPTX 输出路径。The PPTX output path.</param>
        public static void ToPptx(OfficePresentation presentation, string path)
        {
            if (presentation == null) throw new ArgumentNullException("presentation");
            PptxWriter.Write(presentation, path);
        }

        /// <summary>将文档模型按标题层级拆分为演示文稿并保存为 PPTX 文件。Splits the document model into a presentation by heading level and saves it as a PPTX file.</summary>
        /// <param name="document">要转换的文档。The document to convert.</param>
        /// <param name="path">PPTX 输出路径。The PPTX output path.</param>
        public static void DocumentToPptx(OfficeDocument document, string path)
        {
            if (document == null) throw new ArgumentNullException("document");
            PptxWriter.Write(DocumentToPresentationConverter.Convert(document), path);
        }

        /// <summary>将文档模型按标题层级拆分为演示文稿并输出 PPTX 字节数组。Splits the document model into a presentation by heading level and returns the PPTX bytes.</summary>
        /// <param name="document">要转换的文档。The document to convert.</param>
        /// <returns>PPTX 文件字节。The PPTX file bytes.</returns>
        public static byte[] DocumentToPptx(OfficeDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");
            return PptxWriter.Write(DocumentToPresentationConverter.Convert(document));
        }

        /// <summary>将文档模型回写为 Markdown 文本。Writes the document model back to Markdown text.</summary>
        /// <param name="document">要转换的文档。The document to convert.</param>
        /// <returns>生成的 Markdown 文本。The generated Markdown text.</returns>
        public static string ToMarkdown(OfficeDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");
            return Wisdom.Markdown.Markdown.ToMarkdown(document);
        }

        /// <summary>将文档模型回写为 Markdown 文件（UTF-8 无 BOM）。Writes the document model to a Markdown file (UTF-8 without BOM).</summary>
        /// <param name="document">要转换的文档。The document to convert.</param>
        /// <param name="path">Markdown 输出路径。The Markdown output path.</param>
        public static void ToMarkdown(OfficeDocument document, string path)
        {
            if (document == null) throw new ArgumentNullException("document");
            if (path == null) throw new ArgumentNullException("path");
            byte[] text = new UTF8Encoding(false).GetBytes(Wisdom.Markdown.Markdown.ToMarkdown(document));
            File.WriteAllBytes(path, text);
        }
    }
}
