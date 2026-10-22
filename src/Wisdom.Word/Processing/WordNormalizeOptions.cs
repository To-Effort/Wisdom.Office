using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>控制 Word 文档规范化管道的每一步开关。Controls every step of the Word document normalization pipeline.</summary>
    public sealed class WordNormalizeOptions
    {
        // 段落级整理
        /// <summary>删除空段落。Removes empty paragraphs.</summary>
        public bool RemoveEmptyParagraphs { get; set; }
        /// <summary>将过短段落并入相邻段落。Merges overly short paragraphs into their neighbors.</summary>
        public bool MergeShortParagraphs { get; set; }
        /// <summary>判定短段落的字符数阈值。The character count below which a paragraph counts as short.</summary>
        public int ShortParagraphThreshold { get; set; }
        /// <summary>拆分过长段落。Splits overly long paragraphs.</summary>
        public bool SplitLongParagraphs { get; set; }
        /// <summary>判定长段落的字符数阈值。The character count above which a paragraph counts as long.</summary>
        public int LongParagraphThreshold { get; set; }
        /// <summary>统一段落末尾的句末标点风格。Unifies the sentence-ending punctuation style at paragraph ends.</summary>
        public bool UnifyEndingPunctuation { get; set; }
        /// <summary>为缺少句末标点的段落补加标点。Appends ending punctuation to paragraphs that lack it.</summary>
        public bool AddMissingEndingPunctuation { get; set; }
        /// <summary>归一化文本内的空白字符。Normalizes whitespace characters inside the text.</summary>
        public bool NormalizeWhitespace { get; set; }

        // 样式整理
        /// <summary>统一正文与标题样式。Unifies body and heading styles.</summary>
        public bool NormalizeStyles { get; set; }
        /// <summary>将手动格式转换为目标样式格式。Converts manual formatting into style-based formatting.</summary>
        public bool ConvertManualFormatting { get; set; }
        /// <summary>正文中文字体。The East-Asian body font.</summary>
        public string BodyEastAsiaFont { get; set; }
        /// <summary>正文西文字体。The ASCII (Latin) body font.</summary>
        public string BodyAsciiFont { get; set; }
        /// <summary>正文字号（磅）。The body font size in points.</summary>
        public double BodyFontSize { get; set; }
        /// <summary>标题中文字体。The East-Asian heading font.</summary>
        public string HeadingEastAsiaFont { get; set; }
        /// <summary>标题西文字体。The ASCII (Latin) heading font.</summary>
        public string HeadingAsciiFont { get; set; }
        /// <summary>一至九级标题的字号数组（磅）。The heading font-size array for levels one through nine, in points.</summary>
        public double[] HeadingFontSizes { get; set; }
        /// <summary>正文行距倍数。The body line-spacing multiplier.</summary>
        public double BodyLineSpacing { get; set; }
        /// <summary>正文段前间距（磅）。The body space-before value in points.</summary>
        public double BodySpaceBefore { get; set; }
        /// <summary>正文段后间距（磅）。The body space-after value in points.</summary>
        public double BodySpaceAfter { get; set; }
        /// <summary>标题行距倍数。The heading line-spacing multiplier.</summary>
        public double HeadingLineSpacing { get; set; }
        /// <summary>标题段前间距（磅）。The heading space-before value in points.</summary>
        public double HeadingSpaceBefore { get; set; }
        /// <summary>标题段后间距（磅）。The heading space-after value in points.</summary>
        public double HeadingSpaceAfter { get; set; }
        /// <summary>正文首行缩进（字符数）。The first-line indent of body paragraphs measured in characters.</summary>
        public double FirstLineIndentCharacters { get; set; }

        // 标题 & 目录
        /// <summary>将手动加大加粗的伪标题升级为真标题。Promotes manually enlarged and bolded pseudo-headings into real headings.</summary>
        public bool PromoteFakeHeadings { get; set; }
        /// <summary>判定伪标题的最小字号（磅）。The minimum font size used to detect fake headings, in points.</summary>
        public double FakeHeadingMinFontSize { get; set; }
        /// <summary>修正标题级别的跳级。Fixes skipped heading levels.</summary>
        public bool FixHeadingLevelSkips { get; set; }
        /// <summary>重新生成目录。Regenerates the table of contents.</summary>
        public bool RegenerateTableOfContents { get; set; }
        /// <summary>目录标题文本。The title text of the table of contents.</summary>
        public string TableOfContentsTitle { get; set; }
        /// <summary>目录收录的最大标题级别。The maximum heading level included in the table of contents.</summary>
        public int TableOfContentsMaxLevel { get; set; }
        /// <summary>目录结束后插入分页符。Adds a page break after the table of contents.</summary>
        public bool AddPageBreakAfterTableOfContents { get; set; }

        // 表格整理
        /// <summary>统一表格格式。Unifies table formatting.</summary>
        public bool NormalizeTables { get; set; }
        /// <summary>统一表格整体对齐方式。Unifies the overall table alignment.</summary>
        public bool UnifyTableAlignment { get; set; }
        /// <summary>表格目标对齐方式。The target alignment mode for tables.</summary>
        public TableAlignment TableAlignment { get; set; }
        /// <summary>删除整行为空的表格行。Removes fully empty table rows.</summary>
        public bool RemoveEmptyTableRows { get; set; }
        /// <summary>删除整列为空的表格列。Removes fully empty table columns.</summary>
        public bool RemoveEmptyTableColumns { get; set; }
        /// <summary>自动调整表格列宽。Autofits table column widths.</summary>
        public bool AutofitTables { get; set; }
        /// <summary>让表头行跨页重复。Repeats the table header row on each page.</summary>
        public bool RepeatTableHeaderRows { get; set; }

        // 图片整理
        /// <summary>统一图片排版。Unifies image layout.</summary>
        public bool NormalizeImages { get; set; }
        /// <summary>将图片居中。Centers standalone images.</summary>
        public bool CenterImages { get; set; }
        /// <summary>图片最大宽度（磅，0 表示按版心宽）。The maximum image width in points (0 uses the content width).</summary>
        public double MaxImageWidth { get; set; }
        /// <summary>删除极小图片。Removes tiny images at or below the pixel threshold.</summary>
        public bool RemoveTinyImages { get; set; }
        /// <summary>判定极小图片的最大像素边长。The maximum pixel edge size considered tiny.</summary>
        public int TinyImageMaxPixels { get; set; }

        /// <summary>默认构造函数，按中文公文常用值初始化各开关与尺寸。Default constructor initializing every switch and size with common Chinese document defaults.</summary>
        public WordNormalizeOptions()
        {
            RemoveEmptyParagraphs = true;
            MergeShortParagraphs = true;
            ShortParagraphThreshold = 18;
            SplitLongParagraphs = true;
            LongParagraphThreshold = 220;
            UnifyEndingPunctuation = true;
            AddMissingEndingPunctuation = false;
            NormalizeWhitespace = true;

            NormalizeStyles = true;
            ConvertManualFormatting = true;
            BodyEastAsiaFont = "宋体";
            BodyAsciiFont = "Times New Roman";
            BodyFontSize = 12;
            HeadingEastAsiaFont = "黑体";
            HeadingAsciiFont = "Arial";
            HeadingFontSizes = new double[] { 16, 15, 14, 13, 12, 12, 12, 12, 12 };
            BodyLineSpacing = 1.5;
            BodySpaceBefore = 0;
            BodySpaceAfter = 0;
            HeadingLineSpacing = 1.5;
            HeadingSpaceBefore = 12;
            HeadingSpaceAfter = 6;
            FirstLineIndentCharacters = 2;

            PromoteFakeHeadings = true;
            FakeHeadingMinFontSize = 14;
            FixHeadingLevelSkips = true;
            RegenerateTableOfContents = true;
            TableOfContentsTitle = "目录";
            TableOfContentsMaxLevel = 3;
            AddPageBreakAfterTableOfContents = true;

            NormalizeTables = true;
            UnifyTableAlignment = true;
            TableAlignment = TableAlignment.Center;
            RemoveEmptyTableRows = true;
            RemoveEmptyTableColumns = true;
            AutofitTables = true;
            RepeatTableHeaderRows = true;

            NormalizeImages = true;
            CenterImages = true;
            MaxImageWidth = 0;
            RemoveTinyImages = true;
            TinyImageMaxPixels = 1;
        }

        /// <summary>
        /// 返回指定标题级别应使用的字号。
        /// Returns the font size that the given heading level should use.
        /// </summary>
        /// <param name="level">标题级别，从 1 开始。The heading level, starting at 1.</param>
        /// <returns>标题字号（磅）。The heading font size in points.</returns>
        public double GetHeadingFontSize(int level)
        {
            if (HeadingFontSizes == null || HeadingFontSizes.Length == 0) return BodyFontSize + 4;
            int index = level < 1 ? 0 : level - 1;
            if (index >= HeadingFontSizes.Length) index = HeadingFontSizes.Length - 1;
            return HeadingFontSizes[index];
        }
    }
}