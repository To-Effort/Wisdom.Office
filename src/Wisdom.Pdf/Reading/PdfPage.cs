using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Wisdom.Pdf
{
    /// <summary>PDF 文档中的一页。A single page within a PDF document.</summary>
    /// <remarks>
    /// 页面尺寸取自 MediaBox 的宽高差值；缺失或非法时退回 US Letter（612×792）。文本与表格为按需懒解析，同一页只解析一次。
    /// Page size is taken from the width/height delta of the MediaBox; when it is missing or invalid the
    /// code falls back to US Letter (612x792). Text and tables are parsed lazily on demand, and each page
    /// is only extracted once.
    /// </remarks>
    public sealed class PdfPage
    {
        private readonly PdfDocument _document;
        private readonly PdfDictionary _dictionary;
        private readonly PdfDictionary _resources;
        /// <summary>已解析的文本行，null 表示尚未提取。Extracted text lines; null means extraction has not run yet.</summary>
        private List<PdfTextLine> _lines;
        /// <summary>已解析的表格，null 表示尚未提取。Extracted tables; null means extraction has not run yet.</summary>
        private List<PdfTable> _tables;

        /// <summary>
        /// 由文档读取器创建页面对象。
        /// Creates a page object, invoked by the document reader.
        /// </summary>
        /// <param name="document">所属文档，用于解析引用与解码流。The owning document, used to resolve references and decode streams.</param>
        /// <param name="dictionary">页面字典 /Page。The page dictionary (/Page).</param>
        /// <param name="resources">继承而来的资源字典，null 时以空字典代替。The inherited resource dictionary; an empty dictionary is substituted when null.</param>
        /// <param name="mediaBox">媒体框数组（可能仍是引用），用于计算页面尺寸。The MediaBox array (possibly still a reference), used to compute page size.</param>
        /// <param name="rotation">页面 /Rotate 值，会被规整到 0/90/180/270。The raw /Rotate value, normalized to 0/90/180/270.</param>
        internal PdfPage(PdfDocument document, PdfDictionary dictionary, PdfDictionary resources, PdfObject mediaBox, int rotation)
        {
            _document = document;
            _dictionary = dictionary;
            _resources = resources ?? new PdfDictionary();
            Rotation = NormalizeRotation(rotation);
            Width = 612;
            Height = 792;

            PdfArray box = document.Resolve(mediaBox) as PdfArray;
            if (box != null && box.Count >= 4)
            {
                PdfNumber x1 = document.Resolve(box[0]) as PdfNumber;
                PdfNumber y1 = document.Resolve(box[1]) as PdfNumber;
                PdfNumber x2 = document.Resolve(box[2]) as PdfNumber;
                PdfNumber y2 = document.Resolve(box[3]) as PdfNumber;
                if (x1 != null && y1 != null && x2 != null && y2 != null)
                {
                    Width = Math.Abs(x2.Value - x1.Value);
                    Height = Math.Abs(y2.Value - y1.Value);
                }
            }
            if (Width <= 0) Width = 612;
            if (Height <= 0) Height = 792;
        }

        /// <summary>页码，从 1 开始。One-based page number.</summary>
        public int Number { get; internal set; }
        /// <summary>不考虑旋转的页面宽度（PDF 用户单位）。Page width in PDF user units, ignoring rotation.</summary>
        public double Width { get; private set; }
        /// <summary>不考虑旋转的页面高度（PDF 用户单位）。Page height in PDF user units, ignoring rotation.</summary>
        public double Height { get; private set; }
        /// <summary>规整后的页面旋转角度，只取 0、90、180、270。The normalized page rotation; only 0, 90, 180 or 270.</summary>
        public int Rotation { get; private set; }
        /// <summary>所属文档。The owning document.</summary>
        public PdfDocument Document { get { return _document; } }

        /// <summary>页面字典，供渲染与字体解析读取扩展项。The page dictionary; rendering and font resolution read extra entries from it.</summary>
        internal PdfDictionary Dictionary { get { return _dictionary; } }
        /// <summary>页面资源字典（已沿继承链合并）。The page resource dictionary, already merged along the inheritance chain.</summary>
        internal PdfDictionary Resources { get { return _resources; } }

        /// <summary>
        /// 取得本页全部内容的解码字节。
        /// Returns the decoded bytes of all content streams of this page.
        /// </summary>
        /// <returns>合并后的内容流；Contents 为数组时各段以换行分隔拼接，无内容时返回空数组。The merged content stream; when Contents is an array the segments are joined with newline separators, and an empty array is returned when there is no content.</returns>
        public byte[] GetContentBytes()
        {
            PdfObject contents = _document.Resolve(_dictionary.Get("Contents"));
            PdfStream stream = contents as PdfStream;
            if (stream != null) return _document.DecodeStream(stream);

            PdfArray array = contents as PdfArray;
            if (array == null) return new byte[0];

            using (MemoryStream memory = new MemoryStream())
            {
                for (int i = 0; i < array.Count; i++)
                {
                    PdfStream item = _document.Resolve(array[i]) as PdfStream;
                    if (item == null) continue;
                    byte[] data = _document.DecodeStream(item);
                    memory.Write(data, 0, data.Length);
                    memory.WriteByte(10);
                }
                return memory.ToArray();
            }
        }

        /// <summary>
        /// 提取本页的纯文本。
        /// Extracts the plain text of this page.
        /// </summary>
        /// <returns>各行文本以换行符连接的结果。The line texts joined by newline characters.</returns>
        public string ExtractText()
        {
            EnsureExtracted();
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < _lines.Count; i++)
            {
                if (i > 0) builder.Append('\n');
                builder.Append(_lines[i].Text);
            }
            return builder.ToString();
        }

        /// <summary>
        /// 取得本页按基线归组后的文本行。
        /// Returns this page's text lines grouped by baseline.
        /// </summary>
        /// <returns>文本行列表，顺序近似阅读顺序。The list of text lines, ordered approximately by reading order.</returns>
        public IReadOnlyList<PdfTextLine> GetTextLines()
        {
            EnsureExtracted();
            return _lines;
        }

        /// <summary>
        /// 取得本页识别出的表格。
        /// Returns the tables recognized on this page.
        /// </summary>
        /// <returns>表格列表；未识别到表格时为空列表。The table list; an empty list when no table is recognized.</returns>
        public IReadOnlyList<PdfTable> GetTables()
        {
            EnsureExtracted();
            return _tables;
        }

        /// <summary>
        /// 由提取器回填解析结果；null 会被替换为空列表，从而让懒加载标记为已完成。
        /// Fills in the extraction results from the extractor; null is replaced with an empty list so the
        /// lazy-loading flag is considered complete.
        /// </summary>
        /// <param name="lines">文本行。The text lines.</param>
        /// <param name="tables">表格。The tables.</param>
        internal void SetExtraction(List<PdfTextLine> lines, List<PdfTable> tables)
        {
            _lines = lines ?? new List<PdfTextLine>();
            _tables = tables ?? new List<PdfTable>();
        }

        /// <summary>
        /// 首次访问文本或表格时触发整页提取，后续调用直接复用缓存。
        /// Triggers full-page extraction on the first access to text or tables; later calls reuse the cache.
        /// </summary>
        private void EnsureExtracted()
        {
            if (_lines != null && _tables != null) return;
            PdfPageExtractor.Extract(this);
            if (_lines == null) _lines = new List<PdfTextLine>();
            if (_tables == null) _tables = new List<PdfTable>();
        }

        /// <summary>
        /// 把任意旋转角归整到最接近的 90 度倍数，非法角度不会导致渲染方向错乱。
        /// Snaps an arbitrary rotation angle to the nearest multiple of 90 degrees so that invalid angles
        /// cannot scramble the rendering orientation.
        /// </summary>
        /// <param name="rotation">原始 /Rotate 值。The raw /Rotate value.</param>
        /// <returns>0、90、180 或 270。One of 0, 90, 180 or 270.</returns>
        private static int NormalizeRotation(int rotation)
        {
            int value = rotation % 360;
            if (value < 0) value += 360;
            if (value < 45 || value >= 315) return 0;
            if (value < 135) return 90;
            if (value < 225) return 180;
            return 270;
        }
    }
}
