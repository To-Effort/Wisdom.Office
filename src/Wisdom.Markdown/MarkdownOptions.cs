using System;

namespace Wisdom.Markdown
{
    /// <summary>
    /// Markdown 解析与渲染选项。
    /// Options controlling Markdown parsing and rendering.
    /// </summary>
    public sealed class MarkdownOptions
    {
        /// <summary>
        /// 单个换行是否视为硬换行（true 时每行都强制换行，类似 GFM break_lines）。
        /// Whether a single newline is treated as a hard line break (like GFM break-lines when true).
        /// </summary>
        public bool HardLineBreaks { get; set; }

        /// <summary>
        /// 是否解析 GFM 管道表格。
        /// Whether to parse GFM pipe tables.
        /// </summary>
        public bool EnableTables { get; set; }

        /// <summary>
        /// 是否解析任务列表复选框（<c>- [ ]</c> / <c>- [x]</c>）。
        /// Whether to parse task-list checkboxes (<c>- [ ]</c> / <c>- [x]</c>).
        /// </summary>
        public bool EnableTaskLists { get; set; }

        /// <summary>
        /// 是否将原始 HTML 行透传输出（false 时按纯文本保留标签）。
        /// Whether to pass raw HTML lines through verbatim (when false, tags are kept as plain text).
        /// </summary>
        public bool EnableRawHtml { get; set; }

        /// <summary>
        /// 代码块与行内代码使用的等宽字体名。
        /// The monospaced font family used by code blocks and inline code.
        /// </summary>
        public string CodeFontFamily { get; set; }

        /// <summary>
        /// 正文使用的字体名；null 表示沿用文档默认字体。
        /// The body font family; null means the document default font is kept.
        /// </summary>
        public string BodyFontFamily { get; set; }

        /// <summary>
        /// 正文字号基准，单位为磅。
        /// The base body font size, in points.
        /// </summary>
        public double BaseFontSize { get; set; }

        /// <summary>
        /// 是否去除行尾多余空白（硬换行依赖的两个尾部空格除外）。
        /// Whether to trim trailing whitespace (except the two trailing spaces that signal a hard break).
        /// </summary>
        public bool TrimTrailingWhitespace { get; set; }

        /// <summary>
        /// 本地图片的基准目录：仅当该属性非空时解析器才会读盘，且只读取该目录内的文件；
        /// 图片地址中的绝对路径、UNC 路径与 ".." 逃逸都会被拒绝。默认 null，即完全不访问本地文件系统。
        /// The base directory for local images: the parser touches the file system only when this is non-null,
        /// and then only reads files inside that directory. Absolute paths, UNC paths and ".." escapes in the
        /// image source are rejected. The default is null, meaning the local file system is never accessed.
        /// </summary>
        public string ImageBaseDirectory { get; set; }

        /// <summary>
        /// 单个本地图片允许读取的最大字节数，超过即降级为替代文本，默认 16 MB。
        /// The largest number of bytes a single local image may occupy; larger files degrade to the alt text. Default 16 MB.
        /// </summary>
        public int MaxImageBytes { get; set; }

        /// <summary>
        /// 默认构造：启用表格与任务列表、禁用原始 HTML、等宽字体 Consolas、正文沿用默认字体、基准字号 11 磅、去除行尾空白、禁用本地图片读取、图片上限 16 MB。
        /// Default constructor: tables and task lists enabled, raw HTML disabled, Consolas code font, document-default body font,
        /// 11-point base size, trailing whitespace trimmed, local image loading disabled, 16 MB image cap.
        /// </summary>
        public MarkdownOptions()
        {
            HardLineBreaks = false;
            EnableTables = true;
            EnableTaskLists = true;
            EnableRawHtml = false;
            CodeFontFamily = "Consolas";
            BodyFontFamily = null;
            BaseFontSize = 11.0;
            TrimTrailingWhitespace = true;
            ImageBaseDirectory = null;
            MaxImageBytes = DefaultMaxImageBytes;
        }

        /// <summary>默认图片读取上限（16 MB）。The default image read cap (16 MB).</summary>
        internal const int DefaultMaxImageBytes = 16 * 1024 * 1024;

        /// <summary>
        /// 返回当前选项的副本。
        /// Returns a copy of the current options.
        /// </summary>
        /// <returns>
        /// 属性相同的新实例。A new instance with the same properties.
        /// </returns>
        public MarkdownOptions Clone()
        {
            MarkdownOptions copy = new MarkdownOptions();
            copy.HardLineBreaks = HardLineBreaks;
            copy.EnableTables = EnableTables;
            copy.EnableTaskLists = EnableTaskLists;
            copy.EnableRawHtml = EnableRawHtml;
            copy.CodeFontFamily = CodeFontFamily;
            copy.BodyFontFamily = BodyFontFamily;
            copy.BaseFontSize = BaseFontSize;
            copy.TrimTrailingWhitespace = TrimTrailingWhitespace;
            copy.ImageBaseDirectory = ImageBaseDirectory;
            copy.MaxImageBytes = MaxImageBytes;
            return copy;
        }
    }
}
