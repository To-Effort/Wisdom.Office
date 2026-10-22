using System;
using System.Collections.Generic;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Pdf
{
    /// <summary>排版前的单个字形字符，携带字体、字号与宽度等测量结果。A single shaped character before layout, carrying font, size and measured width.</summary>
    internal sealed class ShapedChar
    {
        /// <summary>Unicode 码点。The Unicode code point.</summary>
        public int CodePoint;
        /// <summary>经字体回退选定后实际使用的字体。The font actually selected after fallback resolution.</summary>
        public PdfFont Font;
        /// <summary>生效字号（上下标已折算）。The effective font size (scripts already scaled).</summary>
        public double FontSize;
        /// <summary>该字符的文本次级格式。The character-level text formatting.</summary>
        public TextFormat Format;
        /// <summary>测量得到的前进宽度（点）。The measured advance width in points.</summary>
        public double Width;
        /// <summary>基线偏移量，用于上标/下标。The baseline shift applied for superscript/subscript.</summary>
        public double BaselineShift;
        /// <summary>是否为换行符。Whether this is a newline.</summary>
        public bool IsNewline;
        /// <summary>是否为空格。Whether this is a space.</summary>
        public bool IsSpace;
        /// <summary>是否为 CJK 字符，可逐字断行。Whether this is a CJK character, breakable per glyph.</summary>
        public bool IsCjk;

        /// <summary>
        /// 实际可绘制的码点：控制字符、孤立代理项与越界值退化为 '?'。
        /// The code point actually drawn; control characters, lone surrogates and out-of-range values
        /// degrade to '?'.
        /// </summary>
        public int DrawableCodePoint
        {
            get
            {
                if (CodePoint < 0x20 || CodePoint > 0x10FFFF) return '?';
                if (CodePoint >= 0xD800 && CodePoint <= 0xDFFF) return '?';
                return CodePoint;
            }
        }

        /// <summary>
        /// 把可绘制码点按 UTF-16 追加到构建器，避免逐字生成临时字符串。
        /// Appends the drawable code point as UTF-16 straight into the builder instead of holding a
        /// per-character string.
        /// </summary>
        /// <param name="builder">目标构建器。The target builder.</param>
        public void AppendUtf16(StringBuilder builder)
        {
            int codePoint = DrawableCodePoint;
            if (codePoint <= 0xFFFF)
            {
                builder.Append((char)codePoint);
                return;
            }
            int value = codePoint - 0x10000;
            builder.Append((char)(0xD800 + (value >> 10)));
            builder.Append((char)(0xDC00 + (value & 0x3FF)));
        }
    }

    /// <summary>断行的最小单位：一个不可再拆的词、空格、换行、CJK 单字或图像。The atomic unit of line breaking: an unsplittable word, space, newline, single CJK character or image.</summary>
    internal sealed class LayoutAtom
    {
        /// <summary>组成本原子的字形字符序列。The shaped characters forming this atom.</summary>
        public List<ShapedChar> Chars = new List<ShapedChar>();
        /// <summary>原子为图像时承载的图像，否则为 null。The image carried by an image atom, otherwise null.</summary>
        public OfficeImage Image;
        /// <summary>原子的总宽度（点）。The total width of the atom in points.</summary>
        public double Width;
        /// <summary>图像的排版宽度。The laid-out width of the image.</summary>
        public double ImageWidth;
        /// <summary>图像的排版高度。The laid-out height of the image.</summary>
        public double ImageHeight;
        /// <summary>是否为换行原子。Whether this is a newline atom.</summary>
        public bool IsNewline;
        /// <summary>是否为空格原子。Whether this is a space atom.</summary>
        public bool IsSpace;

        /// <summary>是否为图像原子。Whether this atom is an image atom.</summary>
        public bool IsImage { get { return Image != null; } }

        /// <summary>
        /// 累加各字符宽度得到原子总宽。
        /// Sums the character advances to compute the atom's total width.
        /// </summary>
        public void Measure()
        {
            double width = 0;
            for (int i = 0; i < Chars.Count; i++)
            {
                width += Chars[i].Width;
            }
            Width = width;
        }
    }

    /// <summary>段落后已断好的一行文本。One wrapped line of a paragraph.</summary>
    internal sealed class PdfLayoutLine
    {
        /// <summary>组成该行的原子序列。The atoms forming this line.</summary>
        public List<LayoutAtom> Atoms = new List<LayoutAtom>();
        /// <summary>行内容宽度（点）。The content width of the line in points.</summary>
        public double Width;
        /// <summary>行高，含行距放大。The line height, including line-spacing enlargement.</summary>
        public double Height;
        /// <summary>基线以上的 ascent，决定首字符落点。The ascent above the baseline, deciding where the first glyph lands.</summary>
        public double Ascent;
        /// <summary>该行可用的排版宽度。The available layout width for this line.</summary>
        public double AvailableWidth;
        /// <summary>该行的左缩进。The left indent of this line.</summary>
        public double LeftIndent;
        /// <summary>是否为不含任何原子的空行。Whether this is an empty line without atoms.</summary>
        public bool IsEmpty { get { return Atoms.Count == 0; } }
    }

    /// <summary>把段落内联元素切分为可断行原子的整形器。The shaper that splits paragraph inlines into line-breakable atoms.</summary>
    /// <remarks>
    /// 处理制表符展开、上下标缩放、换行归一、CJK 逐字断行与图像尺寸推导；连续非空白拉丁字符被聚合为
    /// 不可拆的词原子。
    /// It expands tabs, scales superscripts/subscripts, normalizes newlines, breaks CJK text per character
    /// and derives image sizes; consecutive non-space Latin characters are grouped into unsplittable word
    /// atoms.
    /// </remarks>
    internal static class InlineShaper
    {
        /// <summary>
        /// 将段落的内联内容原子化为排版序列。
        /// Atomizes the inline content of a paragraph into a layout sequence.
        /// </summary>
        /// <param name="paragraph">待整形段落。The paragraph to shape.</param>
        /// <param name="defaultFormat">段落默认文字格式。The default text format of the paragraph.</param>
        /// <param name="fonts">字体集合，负责码点回退。The font set providing code-point fallback.</param>
        /// <param name="maxWidth">可用最大宽度，用于等比缩放超宽图像。The maximum available width, used to scale down over-wide images proportionally.</param>
        /// <returns>原子序列，null 段落返回空列表。The atom sequence; a null paragraph yields an empty list.</returns>
        public static List<LayoutAtom> Atomize(Paragraph paragraph, TextFormat defaultFormat, PdfFontSet fonts, double maxWidth)
        {
            List<LayoutAtom> atoms = new List<LayoutAtom>();
            List<ShapedChar> word = new List<ShapedChar>();
            if (paragraph == null) return atoms;
            if (defaultFormat == null) defaultFormat = new TextFormat();

            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                string inlineText = null;
                TextFormat inlineFormat = null;
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run != null)
                {
                    inlineText = run.Text;
                    inlineFormat = run.Format;
                }
                else
                {
                    FieldRun field = paragraph.Inlines[i] as FieldRun;
                    if (field != null)
                    {
                        inlineText = field.DisplayText;
                        inlineFormat = field.Format;
                    }
                }
                if (inlineText != null)
                {
                    TextFormat format = inlineFormat ?? defaultFormat;
                    if (format.FontSize <= 0) format.FontSize = defaultFormat.FontSize;
                    string text = inlineText.Replace("\r\n", "\n").Replace('\r', '\n');
                    bool superscript = format.Superscript && !format.Subscript;
                    bool subscript = format.Subscript && !format.Superscript;
                    double effectiveSize = format.FontSize;
                    double shift = 0;
                    if (superscript)
                    {
                        effectiveSize = format.FontSize * 0.65;
                        shift = -format.FontSize * 0.33;
                    }
                    else if (subscript)
                    {
                        effectiveSize = format.FontSize * 0.65;
                        shift = format.FontSize * 0.16;
                    }

                    foreach (int codePoint in UnicodeUtil.EnumerateCodePoints(text))
                    {
                        if (codePoint == '\t')
                        {
                            FlushWord(word, atoms);
                            string spaces = "    ";
                            for (int s = 0; s < spaces.Length; s++)
                            {
                                ShapedChar spaceChar = CreateChar(' ', format, fonts, effectiveSize, shift);
                                LayoutAtom spaceAtom = new LayoutAtom();
                                spaceAtom.IsSpace = true;
                                spaceAtom.Chars.Add(spaceChar);
                                spaceAtom.Measure();
                                atoms.Add(spaceAtom);
                            }
                            continue;
                        }

                        ShapedChar shaped = CreateChar(codePoint, format, fonts, effectiveSize, shift);
                        if (shaped.IsNewline)
                        {
                            FlushWord(word, atoms);
                            LayoutAtom newlineAtom = new LayoutAtom();
                            newlineAtom.IsNewline = true;
                            newlineAtom.Chars.Add(shaped);
                            newlineAtom.Width = 0;
                            atoms.Add(newlineAtom);
                            continue;
                        }
                        if (shaped.IsSpace)
                        {
                            FlushWord(word, atoms);
                            LayoutAtom spaceAtom = new LayoutAtom();
                            spaceAtom.IsSpace = true;
                            spaceAtom.Chars.Add(shaped);
                            spaceAtom.Measure();
                            atoms.Add(spaceAtom);
                            continue;
                        }
                        if (shaped.IsCjk)
                        {
                            FlushWord(word, atoms);
                            LayoutAtom cjkAtom = new LayoutAtom();
                            cjkAtom.Chars.Add(shaped);
                            cjkAtom.Measure();
                            atoms.Add(cjkAtom);
                            continue;
                        }
                        word.Add(shaped);
                    }
                    FlushWord(word, atoms);
                    continue;
                }

                ImageRun imageRun = paragraph.Inlines[i] as ImageRun;
                if (imageRun != null && imageRun.Image != null)
                {
                    FlushWord(word, atoms);
                    double width = imageRun.Width;
                    double height = imageRun.Height;
                    if (width <= 0 && height <= 0)
                    {
                        width = imageRun.Image.DefaultWidthPoints;
                        height = imageRun.Image.DefaultHeightPoints;
                    }
                    else if (width <= 0)
                    {
                        width = imageRun.Image.PixelWidth > 0
                            ? height * imageRun.Image.PixelWidth / imageRun.Image.PixelHeight
                            : height;
                    }
                    else if (height <= 0)
                    {
                        height = imageRun.Image.PixelHeight > 0
                            ? width * imageRun.Image.PixelHeight / imageRun.Image.PixelWidth
                            : width * 0.75;
                    }
                    if (width > maxWidth && maxWidth > 1)
                    {
                        height *= maxWidth / width;
                        width = maxWidth;
                    }
                    LayoutAtom imageAtom = new LayoutAtom();
                    imageAtom.Image = imageRun.Image;
                    imageAtom.ImageWidth = width;
                    imageAtom.ImageHeight = height;
                    imageAtom.Width = width;
                    atoms.Add(imageAtom);
                }
            }
            FlushWord(word, atoms);
            return atoms;
        }

        /// <summary>
        /// 把累积的字符刷成一个词原子并清空缓冲。
        /// Flushes the accumulated characters into one word atom and clears the buffer.
        /// </summary>
        /// <param name="word">当前词缓冲。The current word buffer.</param>
        /// <param name="atoms">输出原子列表。The output atom list.</param>
        private static void FlushWord(List<ShapedChar> word, List<LayoutAtom> atoms)
        {
            if (word.Count == 0) return;
            LayoutAtom atom = new LayoutAtom();
            atom.Chars.AddRange(word);
            atom.Measure();
            atoms.Add(atom);
            word.Clear();
        }

        /// <summary>
        /// 把段落的项目符号整形为原子序列：符号取段落首个文字运行的格式，末尾补两个空格作为与正文的间距。
        /// Atomizes the paragraph's list marker: the symbol takes the first text run's format and two
        /// trailing spaces provide the gap before the body text.
        /// </summary>
        /// <param name="paragraph">带 <c>ListMarker</c> 的段落。The paragraph carrying a <c>ListMarker</c>.</param>
        /// <param name="defaultFormat">段落默认文字格式。The default text format of the paragraph.</param>
        /// <param name="fonts">字体集合。The font set.</param>
        /// <returns>符号原子；无标记时返回空列表。The marker atoms; an empty list when there is no marker.</returns>
        public static List<LayoutAtom> AtomizeMarker(Paragraph paragraph, TextFormat defaultFormat, PdfFontSet fonts)
        {
            List<LayoutAtom> atoms = new List<LayoutAtom>();
            if (paragraph == null || string.IsNullOrEmpty(paragraph.ListMarker)) return atoms;
            if (defaultFormat == null) defaultFormat = new TextFormat();

            TextFormat format = defaultFormat;
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run != null && !string.IsNullOrEmpty(run.Text) && run.Format != null)
                {
                    format = run.Format;
                    break;
                }
            }
            double size = format.FontSize > 0 ? format.FontSize : defaultFormat.FontSize;
            if (size <= 0) size = 11;

            foreach (int codePoint in UnicodeUtil.EnumerateCodePoints(paragraph.ListMarker))
            {
                ShapedChar shaped = CreateChar(codePoint, format, fonts, size, 0);
                if (shaped.IsNewline) continue;
                LayoutAtom atom = new LayoutAtom();
                atom.IsSpace = shaped.IsSpace;
                atom.Chars.Add(shaped);
                atom.Measure();
                atoms.Add(atom);
            }

            for (int s = 0; s < 2; s++)
            {
                LayoutAtom space = new LayoutAtom();
                space.IsSpace = true;
                space.Chars.Add(CreateChar(' ', format, fonts, size, 0));
                space.Measure();
                atoms.Add(space);
            }
            return atoms;
        }

        /// <summary>
        /// 为码点创建字形字符：解析回退字体、测量前进宽并标记子集用量。
        /// Creates a shaped character for a code point: resolves the fallback font, measures the advance and
        /// marks the glyph as used for subsetting.
        /// </summary>
        /// <param name="codePoint">Unicode 码点。The Unicode code point.</param>
        /// <param name="format">文字格式。The text format.</param>
        /// <param name="fonts">字体集合。The font set.</param>
        /// <param name="size">生效字号。The effective font size.</param>
        /// <param name="shift">基线偏移。The baseline shift.</param>
        /// <returns>填充完成的字形字符。The filled-in shaped character.</returns>
        private static ShapedChar CreateChar(int codePoint, TextFormat format, PdfFontSet fonts, double size, double shift)
        {
            ShapedChar result = new ShapedChar();
            result.CodePoint = codePoint;
            result.Format = format;
            result.FontSize = size;
            result.BaselineShift = shift;
            result.IsNewline = codePoint == '\n';
            result.IsSpace = codePoint == ' ';
            result.IsCjk = FontResolver.IsCjkCodePoint(codePoint);
            if (result.IsNewline)
            {
                result.Font = fonts.GetPrimary(format);
                result.Width = 0;
                return result;
            }
            result.Font = fonts.GetFontForCodePoint(format, codePoint);
            result.Width = result.Font.GetAdvance(codePoint, size);
            result.Font.MarkUsed(result.DrawableCodePoint);
            return result;
        }
    }

    /// <summary>段落断行器，把原子序列按可用宽度贪心分配到各行。The paragraph line breaker, greedily assigning the atom sequence to lines within the available width.</summary>
    /// <remarks>
    /// 首行缩进只作用于第一行；超出整行宽度的长词按字符拆分，超宽图像等比缩放；行尾空格悬挂在行外不计宽度。
    /// The first-line indent applies only to the first line; words wider than the whole line are split per
    /// character, oversized images are scaled down, and trailing spaces hang outside the line without
    /// contributing width.
    /// </remarks>
    internal static class ParagraphLayouter
    {
        /// <summary>
        /// 对段落执行断行，产出各行布局。
        /// Performs line wrapping for a paragraph and produces the laid-out lines.
        /// </summary>
        /// <param name="paragraph">待断行段落。The paragraph to wrap.</param>
        /// <param name="contentWidth">内容区总宽度。The total content width.</param>
        /// <param name="defaultFormat">段落默认格式。The default paragraph format.</param>
        /// <param name="fonts">字体集合。The font set.</param>
        /// <returns>布局行列表，空段落至少返回一条空行。The laid-out lines; an empty paragraph still yields one blank line.</returns>
        public static List<PdfLayoutLine> Wrap(Paragraph paragraph, double contentWidth, TextFormat defaultFormat, PdfFontSet fonts)
        {
            List<PdfLayoutLine> lines = new List<PdfLayoutLine>();
            if (paragraph == null) return lines;
            if (defaultFormat == null) defaultFormat = new TextFormat();

            List<LayoutAtom> markerAtoms = InlineShaper.AtomizeMarker(paragraph, defaultFormat, fonts);
            double hanging = 0;
            for (int m = 0; m < markerAtoms.Count; m++) hanging += markerAtoms[m].Width;

            double left = Math.Max(0, paragraph.LeftIndent);
            double right = Math.Max(0, paragraph.RightIndent);
            double baseAvailable = contentWidth - left - right;

            // 项目符号占用宽度并入左缩进，形成悬挂缩进：续行与正文左端对齐，符号挂在正文左缘之外。
            // The marker width folds into the left indent to form a hanging indent: wrapped lines align
            // with the body text while the symbol hangs outside it.
            if (hanging > 0 && baseAvailable - hanging < 10)
            {
                markerAtoms.Clear();
                hanging = 0;
            }
            left += hanging;
            baseAvailable = contentWidth - left - right;
            if (baseAvailable < 10) baseAvailable = Math.Max(10, contentWidth - left);
            List<LayoutAtom> atoms = new List<LayoutAtom>(markerAtoms);
            atoms.AddRange(InlineShaper.Atomize(paragraph, defaultFormat, fonts, baseAvailable));

            double firstAvailable = baseAvailable + paragraph.FirstLineIndent + hanging;
            double firstLeft = left + paragraph.FirstLineIndent - hanging;
            if (firstLeft < 0) firstLeft = 0;
            if (firstAvailable < 10) firstAvailable = 10;

            List<LayoutAtom> current = new List<LayoutAtom>();
            double currentWidth = 0;
            double available = firstAvailable;
            double lineLeft = firstLeft;
            bool pendingEmptyLine = false;

            Action finalize = null;
            finalize = delegate
            {
                List<LayoutAtom> trimmed = TrimTrailingSpaces(current, ref currentWidth);
                PdfLayoutLine line = BuildLine(trimmed, currentWidth, available, lineLeft, paragraph, defaultFormat, fonts);
                lines.Add(line);
                current.Clear();
                currentWidth = 0;
                available = baseAvailable;
                lineLeft = left;
                pendingEmptyLine = false;
            };

            for (int index = 0; index < atoms.Count; index++)
            {
                LayoutAtom atom = atoms[index];
                if (atom.IsNewline)
                {
                    finalize();
                    pendingEmptyLine = true;
                    continue;
                }
                if (atom.IsSpace && current.Count == 0)
                {
                    continue;
                }

                if (atom.Width <= available - currentWidth + 0.001)
                {
                    current.Add(atom);
                    currentWidth += atom.Width;
                    pendingEmptyLine = false;
                    continue;
                }

                if (atom.IsSpace)
                {
                    if (current.Count > 0)
                    {
                        finalize();
                    }
                    continue;
                }

                if (current.Count > 0)
                {
                    // 避头尾：句号、逗号等后置标点正好落到下一行行首时，允许其悬挂在本行右缘之外，
                    // 避免出现以标点开头的孤行；悬挂量仅一个标点宽，不影响后续行高。
                    // Kinsoku: when a closing punctuation mark would start the next line, let it hang
                    // past the right edge of the current line instead of creating a punctuation-led
                    // orphan line. The overhang is a single glyph and does not affect later line heights.
                    if (IsClosingPunctuation(atom))
                    {
                        current.Add(atom);
                        currentWidth += atom.Width;
                        pendingEmptyLine = false;
                        finalize();
                        continue;
                    }
                    finalize();
                    index--;
                    continue;
                }

                if (atom.IsImage)
                {
                    if (atom.Width > available && available > 1)
                    {
                        double scale = available / atom.Width;
                        atom.ImageWidth *= scale;
                        atom.ImageHeight *= scale;
                        atom.Width = available;
                    }
                    current.Add(atom);
                    currentWidth += atom.Width;
                    pendingEmptyLine = false;
                    finalize();
                    continue;
                }

                if (atom.Chars.Count > 1)
                {
                    LayoutAtom head;
                    LayoutAtom tail;
                    SplitAtom(atom, available, out head, out tail);
                    current.Add(head);
                    currentWidth += head.Width;
                    pendingEmptyLine = false;
                    finalize();
                    if (tail != null && tail.Chars.Count > 0)
                    {
                        atoms.Insert(index + 1, tail);
                    }
                    continue;
                }

                current.Add(atom);
                currentWidth += atom.Width;
                pendingEmptyLine = false;
                finalize();
            }

            if (current.Count > 0)
            {
                finalize();
            }
            if (lines.Count == 0 || pendingEmptyLine)
            {
                lines.Add(BuildLine(new List<LayoutAtom>(), 0, available, lineLeft, paragraph, defaultFormat, fonts));
            }
            return lines;
        }

        /// <summary>
        /// 判断原子是否为不应出现在行首的后置标点（中文句号、逗号、括号等）。
        /// Determines whether an atom is a closing punctuation mark that must not start a line
        /// (Chinese full stop, comma, bracket and similar).
        /// </summary>
        /// <param name="atom">待判断的原子。The atom to test.</param>
        /// <returns>需要悬挂时返回 true。True when the atom should hang on the previous line.</returns>
        private static bool IsClosingPunctuation(LayoutAtom atom)
        {
            if (atom == null || atom.IsSpace || atom.IsNewline || atom.IsImage) return false;
            if (atom.Chars.Count != 1) return false;
            switch (atom.Chars[0].CodePoint)
            {
                case '。':
                case '，':
                case '、':
                case '；':
                case '：':
                case '！':
                case '？':
                case '）':
                case '》':
                case '」':
                case '』':
                case '】':
                case '〕':
                case '”':
                case '’':
                case '…':
                case '～':
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// 去掉行尾空格原子并同步扣减宽度（空格悬挂在行外）。
        /// Removes trailing space atoms and subtracts their width accordingly (spaces hang outside the line).
        /// </summary>
        /// <param name="current">当前行原子。The atoms of the current line.</param>
        /// <param name="width">当前行宽，返回时已扣除悬挂空格。The current line width, reduced by the hanging spaces on return.</param>
        /// <returns>修剪后的原子列表。The trimmed atom list.</returns>
        private static List<LayoutAtom> TrimTrailingSpaces(List<LayoutAtom> current, ref double width)
        {
            List<LayoutAtom> result = new List<LayoutAtom>(current);
            while (result.Count > 0 && result[result.Count - 1].IsSpace)
            {
                width -= result[result.Count - 1].Width;
                result.RemoveAt(result.Count - 1);
            }
            if (width < 0) width = 0;
            return result;
        }

        /// <summary>
        /// 把超宽词原子按字符宽度拆成可放入一行的头部与放入下一行的尾部。
        /// Splits an over-wide word atom per character into a head that fits the line and a tail for the next line.
        /// </summary>
        /// <param name="atom">待拆分的原子。The atom to split.</param>
        /// <param name="maxWidth">行最大可用宽度。The maximum available line width.</param>
        /// <param name="head">头部原子，放入当前行。The head atom placed on the current line.</param>
        /// <param name="tail">尾部原子，可能为空字符列表。The tail atom, which may hold an empty character list.</param>
        private static void SplitAtom(LayoutAtom atom, double maxWidth, out LayoutAtom head, out LayoutAtom tail)
        {
            head = new LayoutAtom();
            tail = new LayoutAtom();
            double width = 0;
            int splitIndex = -1;
            for (int i = 0; i < atom.Chars.Count; i++)
            {
                if (width + atom.Chars[i].Width > maxWidth && i > 0)
                {
                    splitIndex = i;
                    break;
                }
                width += atom.Chars[i].Width;
                splitIndex = i + 1;
            }
            if (splitIndex <= 0) splitIndex = 1;
            for (int i = 0; i < atom.Chars.Count; i++)
            {
                if (i < splitIndex) head.Chars.Add(atom.Chars[i]);
                else tail.Chars.Add(atom.Chars[i]);
            }
            head.Measure();
            tail.Measure();
        }

        /// <summary>
        /// 由原子序列构建布局行：统计最大 ascent/descent，再按行距系数确定行高并居中基线。
        /// Builds a layout line from an atom sequence: it computes the maximum ascent/descent, derives the
        /// line height from the spacing factor and centers the baseline.
        /// </summary>
        /// <param name="atoms">行内原子。The atoms on the line.</param>
        /// <param name="width">行内容宽度。The content width of the line.</param>
        /// <param name="available">可用宽度。The available width.</param>
        /// <param name="leftIndent">左缩进。The left indent.</param>
        /// <param name="paragraph">所属段落，提供行距设置。The owning paragraph, supplying line-spacing settings.</param>
        /// <param name="defaultFormat">默认格式，空行用它兜底度量。The default format, used as a metric fallback for blank lines.</param>
        /// <param name="fonts">字体集合。The font set.</param>
        /// <returns>完成的布局行。The completed layout line.</returns>
        private static PdfLayoutLine BuildLine(List<LayoutAtom> atoms, double width, double available, double leftIndent,
            Paragraph paragraph, TextFormat defaultFormat, PdfFontSet fonts)
        {
            PdfLayoutLine line = new PdfLayoutLine();
            line.Atoms = atoms;
            line.Width = width;
            line.AvailableWidth = available;
            line.LeftIndent = leftIndent;

            double ascent = 0;
            double descent = 0;
            for (int i = 0; i < atoms.Count; i++)
            {
                LayoutAtom atom = atoms[i];
                if (atom.IsImage)
                {
                    if (atom.ImageHeight > ascent) ascent = atom.ImageHeight;
                    continue;
                }
                for (int c = 0; c < atom.Chars.Count; c++)
                {
                    ShapedChar ch = atom.Chars[c];
                    if (ch.IsNewline) continue;
                    double a = ch.Font.GetAscent(ch.FontSize) - ch.BaselineShift;
                    double d = ch.Font.GetDescent(ch.FontSize) + ch.BaselineShift;
                    if (a > ascent) ascent = a;
                    if (d > descent) descent = d;
                }
            }

            if (ascent <= 0 && descent <= 0)
            {
                PdfFont font = fonts.GetPrimary(defaultFormat);
                ascent = font.GetAscent(defaultFormat.FontSize);
                descent = font.GetDescent(defaultFormat.FontSize);
            }

            double natural = ascent + descent;
            double spacing = paragraph.LineSpacing <= 0 ? 1.0 : paragraph.LineSpacing;
            double height = natural * spacing;
            // 空行允许被行距压到自然行高以下：嵌套表格后的占位空段常用极小行距，若仍按自然行高
            // 计算，标签表格会被这些不可见段落撑高。有文字的行保持自然行高，避免上下行叠字。
            // Blank lines may shrink below the natural height: placeholder paragraphs after a nested
            // table often use a tiny line spacing, and forcing natural height would inflate label
            // tables. Lines with glyphs keep the natural height so adjacent lines never overlap.
            if (height < natural && atoms.Count > 0) height = natural;
            double extra = height - natural;
            line.Height = height;
            line.Ascent = ascent + extra / 2;
            return line;
        }
    }

    /// <summary>按内容哈希去重的文档级图像集合。A document-level image collection deduplicated by content hash.</summary>
    /// <remarks>
    /// 相同字节只生成一个 PDF 图像对象并共享资源名，避免重复嵌入。
    /// Identical bytes yield a single PDF image object with a shared resource name, preventing duplicate
    /// embedding.
    /// </remarks>
    internal sealed class PdfImageCollection
    {
        /// <summary>哈希到图像对象的映射。The mapping from content hash to image object.</summary>
        private readonly Dictionary<ulong, PdfImage> _images = new Dictionary<ulong, PdfImage>();
        /// <summary>资源名计数器。The resource-name counter.</summary>
        private int _counter;

        /// <summary>
        /// 取得图像对应的去重 PDF 图像对象。
        /// Gets the deduplicated PDF image object for an image.
        /// </summary>
        /// <param name="image">源图像，null 时返回 null。The source image; null returns null.</param>
        /// <returns>已有的或新建的 PDF 图像。An existing or newly created PDF image.</returns>
        public PdfImage Get(OfficeImage image)
        {
            if (image == null) return null;
            ulong hash = PdfImage.ComputeHash(image.Data);
            hash = hash * 31 + (ulong)image.Data.Length;
            PdfImage result;
            if (!_images.TryGetValue(hash, out result))
            {
                result = new PdfImage(image);
                _counter++;
                result.ResourceName = "Im" + _counter.ToString(System.Globalization.CultureInfo.InvariantCulture);
                _images[hash] = result;
            }
            return result;
        }

        /// <summary>集合内全部去重后的图像。All deduplicated images in the collection.</summary>
        public IEnumerable<PdfImage> All
        {
            get { return _images.Values; }
        }
    }
}
