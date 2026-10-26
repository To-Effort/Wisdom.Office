using System;
using Wisdom.Markdown;
using Md = Wisdom.Markdown.Markdown;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// Markdown 与 CommonMark/Markdig 对齐度的回归测试：字符实体、行内代码、强调贴合规则、块优先级、
    /// 链接引用定义、列表起始编号、标题锚点与代码块空白保真。
    /// Regression tests pinning Markdown behaviour to CommonMark and Markdig: character references, code spans,
    /// emphasis flanking, block precedence, link reference definitions, list start numbers, heading anchors and the
    /// whitespace fidelity of code blocks.
    /// </summary>
    internal static partial class Program
    {
        /// <summary>运行 CommonMark 对齐度回归测试。Runs the CommonMark fidelity regression tests.</summary>
        private static void RunMarkdownFidelityTest()
        {
            RunCharacterReferenceTest();
            RunCodeSpanTest();
            RunEmphasisFlankingTest();
            RunDelimiterRunTest();
            RunBlockPrecedenceTest();
            RunLinkTargetTest();
            RunReferenceDefinitionTest();
            RunListNumberTest();
            RunHeadingAnchorTest();
            RunLineBreakTest();
            RunCodeWhitespaceFidelityTest();
            Console.WriteLine("[mdfidelity] CommonMark fidelity checks ok");
        }

        /// <summary>
        /// 字符实体：数字与命名字先解成字符，写出时再按 HTML 规则转义；认不出的形式原样保留。
        /// Character references: numeric and named forms decode first and are re-escaped on output, while an
        /// unrecognized name stays literal.
        /// </summary>
        private static void RunCharacterReferenceTest()
        {
            RequireContains(Md.ToHtml("&#65;&#x42;&copy;"), "AB\u00A9");
            RequireContains(Md.ToHtml("&amp;"), "&amp;");
            RequireContains(Md.ToHtml("&nbsp;"), "\u00A0");
            RequireContains(Md.ToHtml("&bogus;"), "&amp;bogus;");
            RequireContains(Md.ToHtml("&#0;"), "&amp;#0;");
            RequireContains(Md.ToHtml("&copy"), "&amp;copy");
            // 行内代码里的实体不解码，这正是反引号的用途。References inside a code span stay verbatim - that is the point of backticks.
            RequireContains(Md.ToHtml("`&#65;`"), "<code>&amp;#65;</code>");
        }

        /// <summary>
        /// 行内代码：闭合串必须等长，两侧恰一个空格时各去一格，其余空白原样保留。
        /// Code spans: the closer must be an equally long backtick run, one space on each side is stripped and all
        /// other whitespace survives.
        /// </summary>
        private static void RunCodeSpanTest()
        {
            Require(Md.ToHtml("`a  b`") == "<p><code>a  b</code></p>", "inner double space lost: " + Md.ToHtml("`a  b`"));
            Require(Md.ToHtml("` a `") == "<p><code>a</code></p>", "one space per side not stripped: " + Md.ToHtml("` a `"));
            Require(Md.ToHtml("`` a ` b ``") == "<p><code>a ` b</code></p>", "unequal closer mishandled: " + Md.ToHtml("`` a ` b ``"));
            Require(Md.ToHtml("````  a  ````") == "<p><code> a </code></p>", "four-tick span wrong: " + Md.ToHtml("````  a  ````"));
            RequireContains(Md.ToHtml("``a`"), "``a`");
            Require(Md.ToHtml("` `") == "<p><code> </code></p>", "all-space span was trimmed: " + Md.ToHtml("` `"));
        }

        /// <summary>
        /// 强调的两侧贴合规则：词内下划线不成强调，起始串后必须是内容；相邻同格式的运行共用一对标签。
        /// Emphasis flanking: intraword underscores stay literal and an opener must be followed by content; adjacent
        /// runs of the same format share one pair of tags.
        /// </summary>
        private static void RunEmphasisFlankingTest()
        {
            Require(Md.ToHtml("in_var_iable") == "<p>in_var_iable</p>", "intraword underscore became emphasis: " + Md.ToHtml("in_var_iable"));
            Require(Md.ToHtml("foo_bar_baz") == "<p>foo_bar_baz</p>", "intraword pair still emphasised: " + Md.ToHtml("foo_bar_baz"));
            Require(Md.ToHtml("*a b*") == "<p><em>a b</em></p>", "spaced opener failed: " + Md.ToHtml("*a b*"));
            Require(Md.ToHtml("foo * bar*") == "<p>foo * bar*</p>", "space after delimiter still opened emphasis: " + Md.ToHtml("foo * bar*"));
            Require(Md.ToHtml("_ a_") == "<p>_ a_</p>", "space after underscore opener emphasised: " + Md.ToHtml("_ a_"));
            // 行首的 `* ` 先按无序列表标记吃掉，剩下的 `a*` 只是列表项正文。
            // A leading "* " is claimed by the unordered list marker first, so the trailing asterisk is item text.
            Require(Md.ToHtml("* a*") == "<ul>\n<li>a*</li>\n</ul>", "bullet marker lost to emphasis: " + Md.ToHtml("* a*"));
            Require(Md.ToHtml("**bold *italic* bold**") == "<p><strong>bold <em>italic</em> bold</strong></p>",
                "one bold stretch split into several strong elements: " + Md.ToHtml("**bold *italic* bold**"));
            Require(Md.ToHtml("~~gone~~") == "<p><del>gone</del></p>", "strikeout lost: " + Md.ToHtml("~~gone~~"));
            Require(Md.ToHtml("*a*\n\nb") == "<p><em>a</em></p>\n<p>b</p>", "emphasis across blocks wrong: " + Md.ToHtml("*a*\n\nb"));
        }

        /// <summary>
        /// 相邻分隔符串的配对：单星不抢既能开场又能闭合的双星（*a**b* 的两个星是正文），比起始串短的闭合串
        /// 也闭不了双星（a**b* 整串退回字面），内层成对的星号先拿走自己的闭合位，外层不得越过去抢。
        /// Pairing across adjacent delimiter runs: a single star may not steal a run that can both open and close
        /// (the two stars in *a**b* are content), a closer shorter than its opener cannot pair either (a**b* stays
        /// literal), and an inner matched pair keeps its own closer so the outer opener cannot reach past it.
        /// </summary>
        private static void RunDelimiterRunTest()
        {
            Require(Md.ToHtml("*a**b*") == "<p><em>a**b</em></p>", "inner double star swallowed as closer: " + Md.ToHtml("*a**b*"));
            Require(Md.ToHtml("a**b*") == "<p>a**b*</p>", "short closer paired with double-star opener: " + Md.ToHtml("a**b*"));
            Require(Md.ToHtml("**a *b*") == "<p>**a <em>b</em></p>", "outer opener stole the inner closer: " + Md.ToHtml("**a *b*"));
            Require(Md.ToHtml("**a *b* c**") == "<p><strong>a <em>b</em> c</strong></p>", "nested emphasis broken: " + Md.ToHtml("**a *b* c**"));
            Require(Md.ToHtml("**a***b***") == "<p><strong>a</strong><em>b</em>**</p>", "leftover markers misplaced: " + Md.ToHtml("**a***b***"));
            Require(Md.ToHtml("***a**") == "<p>*<strong>a</strong></p>", "partial pair lost a literal star: " + Md.ToHtml("***a**"));
            Require(Md.ToHtml("**a*") == "<p>*<em>a</em></p>", "line-start double star paired wrong: " + Md.ToHtml("**a*"));
            Require(Md.ToHtml("*a**") == "<p><em>a</em>*</p>", "single opener with double closer wrong: " + Md.ToHtml("*a**"));
            Require(Md.ToHtml("***a*") == "<p>**<em>a</em></p>", "triple opener with single closer wrong: " + Md.ToHtml("***a*"));
            Require(Md.ToHtml("2*3*4 and *not emphasis*?") == "<p>2<em>3</em>4 and <em>not emphasis</em>?</p>",
                "flanking beside digits or punctuation wrong: " + Md.ToHtml("2*3*4 and *not emphasis*?"));

            // 星号只能变成标签，不能凭空消失：剥掉标签后必须还是原文那一串字符。
            // Asterisks may only turn into tags, never vanish: stripping the tags must leave the original characters.
            string stripped = Md.ToHtml("*a**b*")
                .Replace("<p>", string.Empty)
                .Replace("</p>", string.Empty)
                .Replace("<em>", string.Empty)
                .Replace("</em>", string.Empty);
            Require(stripped == "a**b", "emphasis changed the visible text: " + stripped);
        }

        /// <summary>
        /// 块级优先级：分隔线先于列表，缩进代码先于列表且不能打断段落，Setext 标题吃掉整段。
        /// Block precedence: the thematic break is tested before lists, indented code before lists and cannot
        /// interrupt a paragraph, and a Setext heading consumes the whole paragraph.
        /// </summary>
        private static void RunBlockPrecedenceTest()
        {
            string stars = Md.ToHtml("* * *");
            RequireContains(stars, "<hr />");
            Require(stars.IndexOf("<li>", StringComparison.Ordinal) < 0, "* * * became a list: " + stars);
            string dashes = Md.ToHtml("- - -");
            RequireContains(dashes, "<hr />");
            Require(dashes.IndexOf("<li>", StringComparison.Ordinal) < 0, "- - - became a list: " + dashes);

            string indented = Md.ToHtml("    - not a list");
            RequireContains(indented, "<pre><code>- not a list");
            Require(indented.IndexOf("<li>", StringComparison.Ordinal) < 0, "indented marker became a list: " + indented);

            string afterParagraph = Md.ToHtml("para\n    still para");
            Require(afterParagraph.IndexOf("<pre>", StringComparison.Ordinal) < 0, "indented line interrupted a paragraph: " + afterParagraph);

            string setext = Md.ToHtml("line1\nline2\n===");
            RequireContains(setext, "<h1 id=\"line1-line2\">");
            Require(setext.IndexOf("<p>line1", StringComparison.Ordinal) < 0, "setext left a stray paragraph: " + setext);

            string emptyFence = Md.ToHtml("```\n```");
            Require(emptyFence == "<pre><code></code></pre>", "empty fence lost: " + emptyFence);
        }

        /// <summary>
        /// 链接目标合法性：未转义空格使整个结构退回文本，尖括号可包空白，标题与不安全协议各有处置。
        /// Destination rules: an unescaped space makes the construct literal, angle brackets may wrap whitespace,
        /// a title is accepted, and an unsafe scheme is refused.
        /// </summary>
        private static void RunLinkTargetTest()
        {
            Require(Md.ToHtml("[a](/url with space)") == "<p>[a](/url with space)</p>",
                "destination with a space accepted: " + Md.ToHtml("[a](/url with space)"));
            RequireContains(Md.ToHtml("[a](</url with space>)"), "<a href=\"/url with space\">");
            string titled = Md.ToHtml("[a](/u \"the title\")");
            RequireContains(titled, "<a href=\"/u\">");
            Require(titled.IndexOf("the title", StringComparison.Ordinal) < 0, "title leaked into text: " + titled);
            Require(Md.ToHtml("[a]( /url )").IndexOf("<a href=\"/url\">", StringComparison.Ordinal) >= 0,
                "padded destination refused: " + Md.ToHtml("[a]( /url )"));
            string refused = Md.ToHtml("[a](javascript:alert(1))");
            Require(refused.IndexOf("<a href=", StringComparison.Ordinal) < 0, "unsafe scheme emitted: " + refused);
        }

        /// <summary>
        /// 链接引用定义：完整式、简写式与折叠式都能解析，定义行本身不再渲染成文本。
        /// Reference definitions: the full, shortcut and collapsed forms all resolve, and the definition line stops
        /// rendering as text.
        /// </summary>
        private static void RunReferenceDefinitionTest()
        {
            string full = Md.ToHtml("[a][1]\n\n[1]: /url");
            RequireContains(full, "<a href=\"/url\">");
            Require(full.IndexOf("[1]:", StringComparison.Ordinal) < 0, "definition line still rendered: " + full);

            string shortcut = Md.ToHtml("[1]\n\n[1]: /url");
            RequireContains(shortcut, "<a href=\"/url\">");
            RequireContains(shortcut, ">1<");

            string collapsed = Md.ToHtml("[a][]\n\n[a]: /u2");
            RequireContains(collapsed, "<a href=\"/u2\">");

            string caseFolded = Md.ToHtml("[a][REF]\n\n[ref]: /url3");
            RequireContains(caseFolded, "<a href=\"/url3\">");

            string undefined = Md.ToHtml("[a][zz]");
            Require(undefined == "<p>[a][zz]</p>", "undefined reference invented a link: " + undefined);

            string notDefinition = Md.ToHtml("para\n[1]: /url");
            RequireContains(notDefinition, "[1]: /url");

            string document = Md.ToMarkdown(Md.ToDocument("[a][1]\n\n[1]: /url"));
            RequireContains(document, "[a](/url)");

            string image = Md.ToHtml("![alt][pic]\n\n[pic]: /missing.png");
            RequireContains(image, "<p>alt</p>");
        }

        /// <summary>
        /// 列表起始编号：首项不是 1 时输出 start 属性，0 也要输出；无序列表不带。
        /// List start numbers: an ordered list whose first number is not 1 carries the attribute, zero included, and
        /// an unordered list never does.
        /// </summary>
        private static void RunListNumberTest()
        {
            RequireContains(Md.ToHtml("5. five\n6. six"), "<ol start=\"5\">");
            RequireContains(Md.ToHtml("0. zero\n1. one"), "<ol start=\"0\">");
            string plain = Md.ToHtml("1. one\n2. two");
            Require(plain.IndexOf("start=", StringComparison.Ordinal) < 0, "default start emitted: " + plain);
            RequireContains(plain, "<ol>");

            string emptyItem = Md.ToHtml("- x\n-");
            Require(CountOf(emptyItem, "<li>") == 2, "bare marker did not open an item: " + emptyItem);

            string continuation = Md.ToHtml("- a\n  b");
            RequireContains(continuation, "<li>a b</li>");
        }

        /// <summary>
        /// 标题锚点：正文归一化成 id，同名标题依次追加序号。
        /// Heading anchors: the heading text becomes a slug and a repeated slug gains an index.
        /// </summary>
        private static void RunHeadingAnchorTest()
        {
            string once = Md.ToHtml("# Hello World");
            RequireContains(once, "<h1 id=\"hello-world\">");
            string twice = Md.ToHtml("# A\n\n# A");
            RequireContains(twice, "id=\"a\"");
            RequireContains(twice, "id=\"a-1\"");
            string punctuation = Md.ToHtml("# C# & *fancy* name");
            RequireContains(punctuation, "id=\"c-fancy-name\"");
            // 中日韩标题按原文取锚点；Markdig 会把非 ASCII 一律裁成 "section"，那是个用不了的 id。
            // CJK headings keep their own text as the anchor; Markdig collapses non-ASCII to "section", an id nobody can link to.
            Require(Md.ToHtml("# 中文标题").StartsWith("<h1 id=\"中文标题\">", StringComparison.Ordinal),
                "non-ASCII heading anchor lost: " + Md.ToHtml("# 中文标题"));
        }

        /// <summary>
        /// 换行：两个尾部空格与反斜杠都是硬换行，标签后跟真正的换行；HardLineBreaks 让每个换行都成硬换行。
        /// Line breaks: two trailing spaces and a backslash both force a break, each followed by a real newline, and
        /// HardLineBreaks turns every newline into one.
        /// </summary>
        private static void RunLineBreakTest()
        {
            RequireContains(Md.ToHtml("line1  \nline2"), "<br />\nline2");
            RequireContains(Md.ToHtml("line1\\\nline2"), "<br />\nline2");
            MarkdownOptions every = new MarkdownOptions();
            every.HardLineBreaks = true;
            RequireContains(Md.ToHtml("a\nb", every), "<br />\nb");
        }

        /// <summary>
        /// 代码空白保真：块尾换行、块内行尾空格与 Markdown 回写都不能吃掉代码的空白。
        /// Code whitespace fidelity: the closing newline of a block, the trailing spaces of a code line and the
        /// Markdown round-trip must all keep the code's whitespace.
        /// </summary>
        private static void RunCodeWhitespaceFidelityTest()
        {
            Require(Md.ToHtml("```\nx\n```") == "<pre><code>x\n</code></pre>", "block closing newline changed: " + Md.ToHtml("```\nx\n```"));
            string padded = Md.ToHtml("```\nint x;  \n```");
            RequireContains(padded, "int x;  \n");
            string blankInside = Md.ToHtml("```\na\n\nb\n```");
            Require(blankInside == "<pre><code>a\n\nb\n</code></pre>", "blank line inside a block split it: " + blankInside);

            string roundTrip = Md.ToMarkdown(Md.ToDocument("```\nint x;  \n```"));
            RequireContains(roundTrip, "int x;  \n");
            Require(Md.ToHtml(roundTrip) == padded, "code whitespace lost across the round trip: " + Md.ToHtml(roundTrip));
        }

        /// <summary>统计子串出现次数。Counts the occurrences of a substring.</summary>
        /// <param name="text">待统计文本。The text to scan.</param>
        /// <param name="needle">目标子串。The substring to count.</param>
        /// <returns>出现次数。The number of occurrences.</returns>
        private static int CountOf(string text, string needle)
        {
            int count = 0;
            int at = text.IndexOf(needle, StringComparison.Ordinal);
            while (at >= 0)
            {
                count++;
                at = text.IndexOf(needle, at + needle.Length, StringComparison.Ordinal);
            }
            return count;
        }
    }
}
