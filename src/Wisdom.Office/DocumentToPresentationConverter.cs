using System;
using System.Collections.Generic;
using Wisdom.Office.Core;
using Wisdom.Ppt;

namespace Wisdom.Office
{
    /// <summary>
    /// 统一文档模型到演示文稿的转换器：按标题级别切分幻灯片，正文、列表、表格与图片落位到内容占位符，
    /// 用于把 Word 文档或 Markdown 文本一键生成 .pptx。
    /// Converter from the unified document model to a presentation: slides are split by heading level and
    /// body text, lists, tables and pictures are placed into the content placeholder. Used to generate a
    /// .pptx from a Word document or Markdown text in one step.
    /// </summary>
    public static class DocumentToPresentationConverter
    {
        /// <summary>正文占位符的基准字号，单位为磅。Base body font size in points.</summary>
        private const double BodyFontSize = 18.0;
        /// <summary>单页正文允许的最多段落数，超出则续页。Maximum body paragraphs per slide before continuing on a new one.</summary>
        private const int MaxBodyParagraphs = 12;
        /// <summary>表格每行的估算高度，单位为磅。Estimated table row height in points.</summary>
        private const double TableRowHeight = 26.0;

        /// <summary>按 16:9 尺寸转换文档。Converts a document into a 16:9 presentation.</summary>
        /// <param name="document">统一文档模型。The unified document model.</param>
        /// <returns>演示文稿；document 为 null 时返回空演示文稿。The presentation; an empty presentation when document is null.</returns>
        public static OfficePresentation Convert(OfficeDocument document)
        {
            return Convert(document, 960.0, 540.0);
        }

        /// <summary>
        /// 把文档模型转换为演示文稿。一级标题（可选二级标题）开启新页，其余内容按顺序落位，
        /// 正文段落数或纵向空间超出时自动续页。
        /// Converts the document model into a presentation. Level-1 headings (and level-2 when configured)
        /// start a new slide, other content follows in order, and the converter continues onto a new slide
        /// when the paragraph count or the vertical space is exhausted.
        /// </summary>
        /// <param name="document">统一文档模型。The unified document model.</param>
        /// <param name="slideWidth">幻灯片宽度（磅）。Slide width in points.</param>
        /// <param name="slideHeight">幻灯片高度（磅）。Slide height in points.</param>
        /// <returns>演示文稿。The presentation.</returns>
        public static OfficePresentation Convert(OfficeDocument document, double slideWidth, double slideHeight)
        {
            OfficePresentation presentation = new OfficePresentation();
            presentation.SlideWidth = slideWidth > 0 ? slideWidth : 960.0;
            presentation.SlideHeight = slideHeight > 0 ? slideHeight : 540.0;
            if (document == null) return presentation;

            CopyMetadata(document.Metadata, presentation.Metadata);
            if (string.IsNullOrEmpty(presentation.Metadata.Title)) presentation.Metadata.Title = "演示文稿";

            Builder builder = new Builder(presentation);
            for (int s = 0; s < document.Sections.Count; s++)
            {
                List<Block> blocks = document.Sections[s].Blocks;
                for (int i = 0; i < blocks.Count; i++)
                {
                    builder.Append(blocks[i]);
                }
            }
            if (presentation.Slides.Count == 0)
            {
                presentation.AddSlide(SlideLayoutKind.Title);
            }
            return presentation;
        }

        /// <summary>
        /// 顺序消费文档块流的可变状态机：维护当前幻灯片、内容占位符与纵向游标。
        /// Mutable state machine consuming the document block stream: tracks the current slide,
        /// its content placeholder and the vertical cursor.
        /// </summary>
        private sealed class Builder
        {
            /// <summary>目标演示文稿。Target presentation.</summary>
            private readonly OfficePresentation _presentation;
            /// <summary>当前幻灯片。Current slide.</summary>
            private Slide _slide;
            /// <summary>当前内容占位符。Current content placeholder.</summary>
            private TextBoxShape _body;
            /// <summary>表格与图片使用的纵向游标（磅）。Vertical cursor used by tables and pictures, in points.</summary>
            private double _cursor;
            /// <summary>上一页的标题，用于续页标注。Title of the previous slide, used to label continuations.</summary>
            private string _lastTitle = string.Empty;

            /// <summary>绑定演示文稿。Binds the presentation.</summary>
            /// <param name="presentation">目标演示文稿。Target presentation.</param>
            public Builder(OfficePresentation presentation)
            {
                _presentation = presentation;
            }

            /// <summary>消费一个文档块。Consumes one document block.</summary>
            /// <param name="block">文档块。The block.</param>
            public void Append(Block block)
            {
                Paragraph paragraph = block as Paragraph;
                if (paragraph != null)
                {
                    AppendParagraph(paragraph);
                    return;
                }

                Table table = block as Table;
                if (table != null)
                {
                    AppendTable(table);
                    return;
                }

                if (block is PageBreak) NewSlide(string.IsNullOrEmpty(_lastTitle) ? " " : Continuation(_lastTitle));
            }

            /// <summary>
            /// 续页标题：剥掉已有 "（续）" 再补一个，避免同一章连分多页时叠成 "（续）（续）"。
            /// Continuation title: strips an existing "（续）" suffix before appending one, so a third
            /// continuation does not pile up as "（续）（续）".
            /// </summary>
            /// <param name="title">上一页标题。The previous slide's title.</param>
            /// <returns>续页标题。The continuation title.</returns>
            private static string Continuation(string title)
            {
                const string suffix = "（续）";
                string baseTitle = title ?? string.Empty;
                while (baseTitle.Length > suffix.Length && baseTitle.EndsWith(suffix, StringComparison.Ordinal))
                {
                    baseTitle = baseTitle.Substring(0, baseTitle.Length - suffix.Length);
                }
                return baseTitle + suffix;
            }

            /// <summary>按标题级别切分或写入正文。Splits slides on headings or writes body text.</summary>
            /// <param name="paragraph">段落。The paragraph.</param>
            private void AppendParagraph(Paragraph paragraph)
            {
                string text = paragraph.GetText();
                if (paragraph.HeadingLevel == 1)
                {
                    NewSlide(text);
                    return;
                }
                if (paragraph.HeadingLevel == 2)
                {
                    NewSlide(text);
                    return;
                }
                if (text.Length == 0 && paragraph.Inlines.Count == 0) return;

                if (_slide == null) NewSlide(" ");
                if (_body != null && _body.Paragraphs.Count >= MaxBodyParagraphs)
                {
                    NewSlide(Continuation(_lastTitle));
                }

                Paragraph copy = paragraph.Clone();
                StyleBody(copy, copy.HeadingLevel >= 3);
                if (_body != null) _body.Paragraphs.Add(copy);
            }

            /// <summary>写入表格形状。Writes a table shape.</summary>
            /// <param name="table">表格。The table.</param>
            private void AppendTable(Table table)
            {
                if (_slide == null || _body == null) NewSlide(" ");
                double height = Math.Max(TableRowHeight, table.Rows.Count * TableRowHeight);
                EnsureSpace(height);
                TableShape shape = _slide.AddTable(_body.Left, _cursor, _body.Width, height);
                shape.Table = table;
                shape.Name = "表格";
                _cursor += height + 12.0;
            }

            /// <summary>
            /// 统一正文字号与颜色：以演示文稿默认字号为基准，保留粗体、斜体、下划线与颜色，
            /// 三级以上标题按加粗处理。
            /// Normalizes body sizing to the presentation default while keeping bold, italic, underline and
            /// colors; level-3 and deeper headings become bold lead-ins.
            /// </summary>
            /// <param name="paragraph">目标段落。Target paragraph.</param>
            /// <param name="emphasize">是否加粗（三级以上标题）。Whether to bold the run (level-3 and deeper headings).</param>
            private void StyleBody(Paragraph paragraph, bool emphasize)
            {
                double size = BodyFontSize - paragraph.LeftIndent * 0.08;
                if (size < 12.0) size = 12.0;
                for (int i = 0; i < paragraph.Inlines.Count; i++)
                {
                    TextRun run = paragraph.Inlines[i] as TextRun;
                    if (run == null || run.Format == null) continue;
                    run.Format.FontSize = size;
                    if (emphasize) run.Format.Bold = true;
                }
                paragraph.HeadingLevel = 0;
            }

            /// <summary>纵向空间不足时开续页。Starts a continuation slide when the vertical space runs out.</summary>
            /// <param name="needed">需要的最小高度（磅）。Minimum required height in points.</param>
            private void EnsureSpace(double needed)
            {
                if (_body == null) return;
                double limit = _presentation.SlideHeight - _body.Top - 12.0;
                if (_cursor - _body.Top + needed <= limit) return;
                NewSlide(Continuation(_lastTitle));
            }

            /// <summary>新建一页内容幻灯片并重置游标。Starts a new content slide and resets the cursor.</summary>
            /// <param name="title">标题文本。Title text.</param>
            private void NewSlide(string title)
            {
                _lastTitle = title ?? string.Empty;
                _slide = _presentation.AddSlide(SlideLayoutKind.TitleAndContent);
                TextBoxShape titleShape = FindPlaceholder(_slide, PlaceholderKind.Title);
                if (titleShape != null)
                {
                    titleShape.Paragraphs.Clear();
                    Paragraph paragraph = new Paragraph();
                    paragraph.AddText(_lastTitle, _presentation.DefaultTitleFormat.Clone());
                    titleShape.Paragraphs.Add(paragraph);
                }
                _body = FindPlaceholder(_slide, PlaceholderKind.Content);
                if (_body != null)
                {
                    _body.Paragraphs.Clear();
                    _cursor = _body.Top;
                    return;
                }
                _cursor = _presentation.SlideHeight * 0.30;
            }

            /// <summary>按占位符类型查找形状。Finds a shape by placeholder kind.</summary>
            /// <param name="slide">目标幻灯片。Target slide.</param>
            /// <param name="kind">占位符类型。Placeholder kind.</param>
            /// <returns>文本框形状，找不到时返回 null。The text box, or null when absent.</returns>
            private static TextBoxShape FindPlaceholder(Slide slide, PlaceholderKind kind)
            {
                for (int i = 0; i < slide.Shapes.Count; i++)
                {
                    TextBoxShape box = slide.Shapes[i] as TextBoxShape;
                    if (box != null && box.Placeholder == kind) return box;
                }
                return null;
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
