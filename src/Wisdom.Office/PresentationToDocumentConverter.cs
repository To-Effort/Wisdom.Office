using System;
using System.Collections.Generic;
using Wisdom.Office.Core;
using Wisdom.Ppt;

namespace Wisdom.Office
{
    /// <summary>
    /// 演示文稿到统一文档模型的转换器：把每张幻灯片映射为一页（横向、尺寸等于幻灯片），
    /// 形状按 z 序输出为段落、表格与图片，用于把 .pptx 导出为 PDF 或 Word。
    /// Converter from a presentation to the unified document model: each slide becomes one page
    /// (landscape, sized to the slide), and shapes are emitted in z-order as paragraphs, tables and
    /// pictures. Used to export .pptx to PDF or Word.
    /// </summary>
    public static class PresentationToDocumentConverter
    {
        /// <summary>幻灯片页边距，单位为磅。Slide page margin in points.</summary>
        private const double PageMargin = 12.0;

        /// <summary>按默认设置转换（不含备注）。Converts with default settings (notes excluded).</summary>
        /// <param name="presentation">演示文稿。The presentation.</param>
        /// <returns>统一文档模型。The unified document model.</returns>
        public static OfficeDocument Convert(OfficePresentation presentation)
        {
            return Convert(presentation, false);
        }

        /// <summary>
        /// 把演示文稿转换为统一文档模型。
        /// Converts a presentation into the unified document model.
        /// </summary>
        /// <param name="presentation">演示文稿。The presentation.</param>
        /// <param name="includeNotes">是否在每页末尾追加演讲者备注。Whether speaker notes are appended after each slide.</param>
        /// <returns>统一文档模型；presentation 为 null 时返回空文档。The unified document model; an empty document when presentation is null.</returns>
        public static OfficeDocument Convert(OfficePresentation presentation, bool includeNotes)
        {
            OfficeDocument document = new OfficeDocument();
            if (presentation == null) return document;

            CopyMetadata(presentation.Metadata, document.Metadata);
            document.DefaultTextFormat = presentation.GetEffectiveTextFormat().Clone();

            PageSetup setup = new PageSetup();
            setup.Width = presentation.SlideWidth;
            setup.Height = presentation.SlideHeight;
            setup.MarginLeft = PageMargin;
            setup.MarginRight = PageMargin;
            setup.MarginTop = PageMargin;
            setup.MarginBottom = PageMargin;
            setup.HeaderDistance = 6.0;
            setup.FooterDistance = 6.0;

            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                Slide slide = presentation.Slides[i];
                Section section = document.AddSection(setup.Clone());
                // 幻灯片导出必须严格一页幻对应一页：关闭自动续页，超出可印区的内容留在本页（与 PowerPoint 导出裁切一致）。
                // Slide export must be strictly one PDF page per slide: auto page breaks are off so overflow stays on the page (same clipping as PowerPoint's own export).
                section.AllowAutoPageBreak = false;
                List<SlideShape> shapes = OrderedShapes(slide);
                for (int s = 0; s < shapes.Count; s++)
                {
                    AppendShape(section, shapes[s], presentation);
                }
                if (includeNotes && !string.IsNullOrEmpty(slide.NotesText))
                {
                    AppendNotes(section, slide.NotesText);
                }
            }

            if (document.Sections.Count == 0) document.AddSection(setup.Clone());
            return document;
        }

        /// <summary>
        /// 返回按 z 序排列的形状，并把标题占位符提到最前，保证导出文档里标题先出现。
        /// Returns shapes in z-order with the title placeholder moved first so headings lead the exported document.
        /// </summary>
        /// <param name="slide">源幻灯片。The source slide.</param>
        /// <returns>排序后的形状列表。The ordered shapes.</returns>
        private static List<SlideShape> OrderedShapes(Slide slide)
        {
            List<SlideShape> result = new List<SlideShape>();
            for (int i = 0; i < slide.Shapes.Count; i++)
            {
                TextBoxShape box = slide.Shapes[i] as TextBoxShape;
                if (box != null && box.Placeholder == PlaceholderKind.Title)
                {
                    result.Insert(0, box);
                    continue;
                }
                result.Add(slide.Shapes[i]);
            }
            return result;
        }

        /// <summary>把单个形状写入节。Writes one shape into the section.</summary>
        /// <param name="section">目标节。Target section.</param>
        /// <param name="shape">形状。The shape.</param>
        /// <param name="presentation">所属演示文稿，用于取默认字体。Owning presentation, used for default fonts.</param>
        private static void AppendShape(Section section, SlideShape shape, OfficePresentation presentation)
        {
            TableShape table = shape as TableShape;
            if (table != null)
            {
                if (table.Table != null && table.Table.Rows.Count > 0) section.Blocks.Add(table.Table);
                return;
            }

            PictureShape picture = shape as PictureShape;
            if (picture != null)
            {
                if (picture.Image != null)
                {
                    Paragraph paragraph = new Paragraph();
                    paragraph.Alignment = ParagraphAlignment.Center;
                    paragraph.AddImage(picture.Image, picture.Width, picture.Height);
                    section.Blocks.Add(paragraph);
                }
                return;
            }

            TextBoxShape text = shape as TextBoxShape;
            if (text != null)
            {
                AppendText(section, text, presentation);
            }
        }

        /// <summary>
        /// 写入文本框内容：标题占位符按一级标题呈现，其余段落沿用原有对齐与列表标记。
        /// Writes text box content: the title placeholder becomes a level-1 heading, other paragraphs
        /// keep their alignment and list markers.
        /// </summary>
        /// <param name="section">目标节。Target section.</param>
        /// <param name="text">文本框形状。The text box shape.</param>
        /// <param name="presentation">所属演示文稿。Owning presentation.</param>
        private static void AppendText(Section section, TextBoxShape text, OfficePresentation presentation)
        {
            bool isTitle = text.Placeholder == PlaceholderKind.Title;
            for (int i = 0; i < text.Paragraphs.Count; i++)
            {
                Paragraph source = text.Paragraphs[i];
                if (source == null) continue;
                if (source.Inlines.Count == 0 && source.GetText().Length == 0) continue;

                Paragraph paragraph = source.Clone();
                if (isTitle)
                {
                    paragraph.HeadingLevel = 1;
                    paragraph.Alignment = ParagraphAlignment.Left;
                    paragraph.SpaceBefore = 0;
                }
                ApplyDefaultFont(paragraph, isTitle ? presentation.GetEffectiveTitleFormat() : presentation.GetEffectiveTextFormat());
                section.Blocks.Add(paragraph);
            }
        }

        /// <summary>
        /// 为未显式设置字体的文本运行补上演示文稿的默认字体，避免导出时字号忽大忽小。
        /// Fills the presentation default font into runs without an explicit family or size so the
        /// exported document keeps a consistent scale.
        /// </summary>
        /// <param name="paragraph">目标段落。Target paragraph.</param>
        /// <param name="fallback">默认字体格式。Default text format.</param>
        private static void ApplyDefaultFont(Paragraph paragraph, TextFormat fallback)
        {
            if (paragraph == null || fallback == null) return;
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run == null || run.Format == null) continue;
                if (string.IsNullOrEmpty(run.Format.FontFamily))
                {
                    run.Format.FontFamily = fallback.FontFamily;
                    run.Format.AsciiFontFamily = fallback.AsciiFontFamily;
                    run.Format.EastAsiaFontFamily = fallback.EastAsiaFontFamily;
                }
                if (run.Format.FontSize <= 0) run.Format.FontSize = fallback.FontSize;
            }
        }

        /// <summary>追加演讲者备注段落。Appends the speaker notes paragraph.</summary>
        /// <param name="section">目标节。Target section.</param>
        /// <param name="notes">备注文本。Notes text.</param>
        private static void AppendNotes(Section section, string notes)
        {
            string[] lines = notes.Replace("\r\n", "\n").Split('\n');
            TextFormat format = new TextFormat("Microsoft YaHei", 9.0) { Italic = true, Color = new OfficeColor(0x60, 0x60, 0x60) };
            for (int i = 0; i < lines.Length; i++)
            {
                if (string.IsNullOrEmpty(lines[i])) continue;
                Paragraph paragraph = new Paragraph();
                paragraph.AddText((i == 0 ? "备注：" : "      ") + lines[i], format);
                paragraph.SpaceBefore = i == 0 ? 12.0 : 0;
                section.Blocks.Add(paragraph);
            }
        }

        /// <summary>复制元数据。Copies metadata.</summary>
        /// <param name="source">源元数据。Source metadata.</param>
        /// <param name="target">目标元数据。Target metadata.</param>
        private static void CopyMetadata(DocumentMetadata source, DocumentMetadata target)
        {
            if (source == null || target == null) return;
            target.Title = source.Title;
            target.Author = source.Author;
            target.Subject = source.Subject;
            target.Keywords = source.Keywords;
            target.Creator = source.Creator;
            target.Created = source.Created;
            target.Modified = source.Modified;
        }
    }
}
