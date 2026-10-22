using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Pdf
{
    /// <summary>
    /// PDF 文件结构写出器：分配间接对象号，并输出对象体、经典交叉表（xref）、trailer 与 startxref。
    /// PDF file structure writer: assigns indirect object numbers and emits object bodies, a classic
    /// cross-reference table (xref), the trailer and startxref.
    /// </summary>
    /// <remarks>登记接口 <see cref="Add"/> 声明为虚方法，派生类可把对象直接写入已有的对象级文档而非本地表。
    /// <see cref="Add"/> is virtual so a derived writer can register objects into an existing object-level
    /// document instead of the local table.</remarks>
    public class PdfObjectWriter
    {
        /// <summary>内部登记表项，记录对象号及其对应的 PDF 对象。Internal registry entry recording an object number and its PDF object.</summary>
        private sealed class Entry
        {
            /// <summary>分配到的对象号。The assigned object number.</summary>
            public int Number;
            /// <summary>待写出的对象。The object to be written.</summary>
            public PdfObject Object;
        }

        /// <summary>已登记的间接对象列表，写出顺序与登记顺序一致。The registered indirect objects; write order matches registration order.</summary>
        private readonly List<Entry> _entries = new List<Entry>();
        /// <summary>下一个可分配的对象号，从 1 开始递增。The next assignable object number, starting at 1 and incrementing.</summary>
        private int _nextNumber = 1;

        /// <summary>已登记的间接对象数量。The number of registered indirect objects.</summary>
        public int ObjectCount { get { return _entries.Count; } }

        /// <summary>
        /// 登记一个间接对象并分配下一个可用对象号。
        /// Registers an indirect object and assigns it the next available object number.
        /// </summary>
        /// <param name="value">要作为间接对象写出的 PDF 对象。The PDF object to write as an indirect object.</param>
        /// <returns>指向该对象的引用，供其他对象或 trailer 引用。A reference to the object, usable by other objects or the trailer.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="value"/> 为 null 时抛出。Thrown when <paramref name="value"/> is null.</exception>
        public virtual PdfReference Add(PdfObject value)
        {
            if (value == null) throw new ArgumentNullException("value");
            Entry entry = new Entry();
            entry.Number = _nextNumber++;
            entry.Object = value;
            _entries.Add(entry);
            return new PdfReference(entry.Number, 0);
        }

        /// <summary>
        /// 以指定字典登记一个流对象。
        /// Registers a stream object with the given dictionary.
        /// </summary>
        /// <param name="dictionary">流字典（/Type、/Subtype 等）。The stream dictionary (/Type, /Subtype, etc.).</param>
        /// <param name="data">流内容字节。The stream content bytes.</param>
        /// <param name="compress">是否对该流做 Flate 压缩。Whether to Flate-compress the stream.</param>
        /// <returns>指向该流的引用。A reference to the stream.</returns>
        public PdfReference AddStream(PdfDictionary dictionary, byte[] data, bool compress)
        {
            PdfStream stream = new PdfStream(dictionary, data, compress);
            return Add(stream);
        }

        /// <summary>
        /// 登记一个不带预置字典的流对象。
        /// Registers a stream object without a preset dictionary.
        /// </summary>
        /// <param name="data">流内容字节。The stream content bytes.</param>
        /// <param name="compress">是否对该流做 Flate 压缩。Whether to Flate-compress the stream.</param>
        /// <returns>指向该流的引用。A reference to the stream.</returns>
        public PdfReference AddStream(byte[] data, bool compress)
        {
            return Add(new PdfStream(data, compress));
        }

        /// <summary>
        /// 生成完整 PDF 文件字节。
        /// Produces the complete PDF file bytes.
        /// </summary>
        /// <param name="root">目录对象（/Root）引用，不可为 null。The catalog (/Root) reference; must not be null.</param>
        /// <param name="info">信息字典（/Info）引用，可为 null。The information dictionary (/Info) reference; may be null.</param>
        /// <returns>可直接落盘的 PDF 字节数组。A PDF byte array ready to be saved to disk.</returns>
        public byte[] ToArray(PdfReference root, PdfReference info)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                Write(memory, root, info);
                return memory.ToArray();
            }
        }

        /// <summary>
        /// 把文档头部、全部间接对象、交叉表与 trailer 依次写入输出流。
        /// Writes the document header, all indirect objects, the cross-reference table and the trailer to
        /// the output stream, in order.
        /// </summary>
        /// <param name="output">目标流，必须是可寻址的（需要记录对象偏移）。The target stream, which must be seekable (object offsets are recorded).</param>
        /// <param name="root">目录对象引用。The catalog object reference.</param>
        /// <param name="info">信息字典引用，可为 null。The information dictionary reference; may be null.</param>
        /// <exception cref="ArgumentNullException"><paramref name="output"/> 或 <paramref name="root"/> 为 null 时抛出。Thrown when <paramref name="output"/> or <paramref name="root"/> is null.</exception>
        /// <remarks>交叉表条目统一使用代号 0，且所有偏移基于输出流的绝对位置，因此调用方不应在写入前向流中预写内容。
        /// All cross-reference entries use generation 0, and every offset is relative to the absolute stream
        /// position, so callers must not pre-write anything to the stream beforehand.</remarks>
        public void Write(Stream output, PdfReference root, PdfReference info)
        {
            if (output == null) throw new ArgumentNullException("output");
            if (root == null) throw new ArgumentNullException("root");

            List<long> offsets = new List<long>(_entries.Count);
            WriteAscii(output, "%PDF-1.7\n");
            output.Write(new byte[] { 0x25, 0xD0, 0xD3, 0xC4, 0xE5, 0x0A }, 0, 6);

            for (int i = 0; i < _entries.Count; i++)
            {
                Entry entry = _entries[i];
                offsets.Add(output.Position);
                WriteAscii(output, entry.Number.ToString(CultureInfo.InvariantCulture) + " 0 obj\n");
                WriteObject(output, entry.Object);
                WriteAscii(output, "\nendobj\n");
            }

            long xrefOffset = output.Position;
            WriteAscii(output, "xref\n0 " + (_entries.Count + 1).ToString(CultureInfo.InvariantCulture) + "\n");
            WriteAscii(output, "0000000000 65535 f\r\n");
            for (int i = 0; i < offsets.Count; i++)
            {
                WriteAscii(output, offsets[i].ToString("0000000000", CultureInfo.InvariantCulture) + " 00000 n\r\n");
            }

            PdfDictionary trailer = new PdfDictionary();
            trailer.SetNumber("Size", _entries.Count + 1);
            trailer.Set("Root", root);
            if (info != null)
            {
                trailer.Set("Info", info);
            }
            WriteAscii(output, "trailer\n");
            WriteObject(output, trailer);
            WriteAscii(output, "\nstartxref\n" + xrefOffset.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
        }

        /// <summary>
        /// 按对象实际类型递归写出其 PDF 语法文本。
        /// Writes the PDF syntax text of an object recursively based on its concrete type.
        /// </summary>
        /// <param name="output">目标流。The target stream.</param>
        /// <param name="value">待写出的对象，null 时写出关键字 null。The object to write; null emits the keyword null.</param>
        /// <exception cref="NotSupportedException">遇到本写出器不支持的对象类型时抛出。Thrown when an unsupported object type is encountered.</exception>
        /// <remarks>写出流对象时会先克隆字典并按需压缩，避免修改调用方持有的原始字典与字节。
        /// When a stream is written its dictionary is cloned and compressed on demand, so the original
        /// dictionary and bytes held by the caller are never mutated.</remarks>
        private void WriteObject(Stream output, PdfObject value)
        {
            if (value == null)
            {
                WriteAscii(output, "null");
                return;
            }

            PdfNull nullValue = value as PdfNull;
            if (nullValue != null) { WriteAscii(output, "null"); return; }

            PdfBoolean boolean = value as PdfBoolean;
            if (boolean != null) { WriteAscii(output, boolean.Value ? "true" : "false"); return; }

            PdfNumber number = value as PdfNumber;
            if (number != null) { WriteAscii(output, PdfNumberFormat.Format(number.Value)); return; }

            PdfName name = value as PdfName;
            if (name != null) { WriteAscii(output, "/" + PdfNameEscape.Escape(name.Value)); return; }

            PdfString text = value as PdfString;
            if (text != null) { WriteString(output, text); return; }

            PdfReference reference = value as PdfReference;
            if (reference != null)
            {
                WriteAscii(output, reference.Number.ToString(CultureInfo.InvariantCulture)
                    + " " + reference.Generation.ToString(CultureInfo.InvariantCulture) + " R");
                return;
            }

            PdfArray array = value as PdfArray;
            if (array != null)
            {
                WriteAscii(output, "[");
                for (int i = 0; i < array.Count; i++)
                {
                    if (i > 0) WriteAscii(output, " ");
                    WriteObject(output, array[i]);
                }
                WriteAscii(output, "]");
                return;
            }

            PdfDictionary dictionary = value as PdfDictionary;
            if (dictionary != null)
            {
                WriteDictionary(output, dictionary);
                return;
            }

            PdfStream stream = value as PdfStream;
            if (stream != null)
            {
                PdfDictionary dict = stream.Dictionary.Clone();
                byte[] data = stream.Data;
                if (stream.Compress)
                {
                    data = Zlib.Compress(data);
                    dict.Remove("Filter");
                    dict.SetName("Filter", "FlateDecode");
                    dict.Remove("DecodeParms");
                }
                dict.SetNumber("Length", data.Length);
                WriteDictionary(output, dict);
                WriteAscii(output, "\nstream\n");
                output.Write(data, 0, data.Length);
                WriteAscii(output, "\nendstream");
                return;
            }

            throw new NotSupportedException("Unsupported PDF object type: " + value.GetType().FullName);
        }

        /// <summary>
        /// 以 &lt;&lt; &gt;&gt; 形式写出字典，键名经过转义且顺序与字典内部一致。
        /// Writes a dictionary in &lt;&lt; &gt;&gt; form, with escaped key names in the dictionary's internal order.
        /// </summary>
        /// <param name="output">目标流。The target stream.</param>
        /// <param name="dictionary">待写出的字典。The dictionary to write.</param>
        private void WriteDictionary(Stream output, PdfDictionary dictionary)
        {
            WriteAscii(output, "<<");
            bool first = true;
            foreach (KeyValuePair<string, PdfObject> pair in dictionary)
            {
                if (!first) WriteAscii(output, " ");
                first = false;
                WriteAscii(output, "/" + PdfNameEscape.Escape(pair.Key) + " ");
                WriteObject(output, pair.Value);
            }
            WriteAscii(output, ">>");
        }

        /// <summary>
        /// 写出字符串对象：十六进制形式直接转 &lt;&gt;，字面形式则对括号、反斜杠、换行及不可打印字节做转义。
        /// Writes a string object: hex forms become &lt;&gt; directly, while literal forms escape parentheses,
        /// backslashes, newlines and unprintable bytes.
        /// </summary>
        /// <param name="output">目标流。The target stream.</param>
        /// <param name="value">待写出的字符串对象。The string object to write.</param>
        private static void WriteString(Stream output, PdfString value)
        {
            if (value.IsHex)
            {
                StringBuilder builder = new StringBuilder(value.Bytes.Length * 2 + 2);
                builder.Append('<');
                for (int i = 0; i < value.Bytes.Length; i++)
                {
                    builder.Append(value.Bytes[i].ToString("x2", CultureInfo.InvariantCulture));
                }
                builder.Append('>');
                WriteAscii(output, builder.ToString());
                return;
            }

            output.WriteByte((byte)'(');
            for (int i = 0; i < value.Bytes.Length; i++)
            {
                byte b = value.Bytes[i];
                if (b == (byte)'(' || b == (byte)')' || b == (byte)'\\')
                {
                    output.WriteByte((byte)'\\');
                    output.WriteByte(b);
                }
                else if (b == 10)
                {
                    WriteAscii(output, "\\n");
                }
                else if (b == 13)
                {
                    WriteAscii(output, "\\r");
                }
                else if (b == 9)
                {
                    WriteAscii(output, "\\t");
                }
                else if (b < 32 || b > 126)
                {
                    WriteAscii(output, "\\" + Convert.ToString(b, 8).PadLeft(3, '0'));
                }
                else
                {
                    output.WriteByte(b);
                }
            }
            output.WriteByte((byte)')');
        }

        /// <summary>
        /// 以 ASCII 编码写入一段文本，用于所有结构性关键字（PDF 语法要求这部分不含非 ASCII 字节）。
        /// Writes text as ASCII bytes, used for all structural keywords (PDF syntax requires these parts to
        /// contain no non-ASCII bytes).
        /// </summary>
        /// <param name="output">目标流。The target stream.</param>
        /// <param name="text">待写入的文本。The text to write.</param>
        private static void WriteAscii(Stream output, string text)
        {
            byte[] bytes = Encoding.ASCII.GetBytes(text);
            output.Write(bytes, 0, bytes.Length);
        }

        /// <summary>
        /// 把时间格式化为 PDF 日期字符串（D:yyyyMMddHHmmss+hh'mm'），附带本地时区偏移。
        /// Formats a time as a PDF date string (D:yyyyMMddHHmmss+hh'mm') including the local time-zone offset.
        /// </summary>
        /// <param name="value">本地时间。The local time.</param>
        /// <returns>符合 PDF 规范的日期文本。The PDF-conformant date text.</returns>
        public static string FormatPdfDate(DateTime value)
        {
            TimeSpan offset = TimeZoneInfo.Local.GetUtcOffset(value);
            string sign = offset.TotalMinutes >= 0 ? "+" : "-";
            int total = (int)Math.Abs(offset.TotalMinutes);
            return "D:" + value.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture)
                + sign + (total / 60).ToString("00", CultureInfo.InvariantCulture)
                + "'" + (total % 60).ToString("00", CultureInfo.InvariantCulture) + "'";
        }
    }
}
