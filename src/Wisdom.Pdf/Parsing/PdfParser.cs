using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace Wisdom.Pdf
{
    /// <summary>
    /// PDF 词法与对象语法解析器，在一块完整的文件字节上按位置游标推进。
    /// PDF lexical and object-syntax parser advancing a cursor over a complete file byte buffer.
    /// </summary>
    /// <remarks>解析过程容忍不规范的分隔与结尾（例如缺失 endstream 时回退到按标记定位），但不做加解密，交给上层处理。
    /// Parsing tolerates malformed delimiters and endings (for example falling back to token search when endstream is missing), but performs no encryption or decryption, which the caller handles.</remarks>
    internal sealed class PdfParser
    {
        /// <summary>被解析的完整字节缓冲。The complete byte buffer being parsed.</summary>
        private readonly byte[] _data;
        /// <summary>间接对象解析委托，供 /Length 为引用时取真实长度。Resolver delegate for indirect objects, used to fetch the real length when /Length is a reference.</summary>
        private readonly Func<PdfReference, PdfObject> _resolve;

        /// <summary>
        /// 在指定字节缓冲上创建解析器。
        /// Creates a parser over the given byte buffer.
        /// </summary>
        /// <param name="data">待解析的字节，不可为 null。The bytes to parse; must not be null.</param>
        /// <param name="position">起始游标位置。The starting cursor position.</param>
        /// <param name="resolve">间接对象解析委托，可为 null。The indirect-object resolver delegate; may be null.</param>
        /// <exception cref="ArgumentNullException"><paramref name="data"/> 为 null 时抛出。Thrown when <paramref name="data"/> is null.</exception>
        public PdfParser(byte[] data, int position, Func<PdfReference, PdfObject> resolve)
        {
            if (data == null) throw new ArgumentNullException("data");
            _data = data;
            Position = position;
            _resolve = resolve;
        }

        /// <summary>底层字节缓冲，供调用方直接窥探当前字符。The underlying byte buffer, so callers can peek at the current byte directly.</summary>
        public byte[] Data { get { return _data; } }
        /// <summary>当前读取位置（字节偏移），可由调用方保存与回退。The current read position (byte offset), which callers may save and rewind.</summary>
        public int Position { get; set; }
        /// <summary>缓冲总长度。The total length of the buffer.</summary>
        public int Length { get { return _data.Length; } }
        /// <summary>游标是否已到达缓冲末尾。Whether the cursor has reached the end of the buffer.</summary>
        public bool EndOfData { get { return Position >= _data.Length; } }

        /// <summary>
        /// 判断字节是否为 PDF 空白字符（NUL、制表、换行、换页、回车、空格）。
        /// Determines whether a byte is a PDF whitespace character (NUL, tab, line feed, form feed, carriage return, space).
        /// </summary>
        /// <param name="value">待判断的字节。The byte to test.</param>
        /// <returns>是空白返回 true。True when the byte is whitespace.</returns>
        public static bool IsWhitespace(byte value)
        {
            return value == 0 || value == 9 || value == 10 || value == 12 || value == 13 || value == 32;
        }

        /// <summary>
        /// 判断字节是否为 PDF 定界符（空白加括号、尖括号、方括号、花括号、斜杠、百分号）。
        /// Determines whether a byte is a PDF delimiter (whitespace plus parentheses, angle brackets, square brackets, braces, slash, percent).
        /// </summary>
        /// <param name="value">待判断的字节。The byte to test.</param>
        /// <returns>是定界符返回 true。True when the byte is a delimiter.</returns>
        public static bool IsDelimiter(byte value)
        {
            return IsWhitespace(value) || value == (byte)'(' || value == (byte)')' || value == (byte)'<'
                || value == (byte)'>' || value == (byte)'[' || value == (byte)']' || value == (byte)'{'
                || value == (byte)'}' || value == (byte)'/' || value == (byte)'%';
        }

        /// <summary>
        /// 跳过空白与以 % 开始的注释行。
        /// Skips whitespace and comment lines starting with %.
        /// </summary>
        public void SkipWhitespaceAndComments()
        {
            while (Position < _data.Length)
            {
                byte value = _data[Position];
                if (IsWhitespace(value))
                {
                    Position++;
                    continue;
                }
                if (value == (byte)'%')
                {
                    Position++;
                    while (Position < _data.Length && _data[Position] != 10 && _data[Position] != 13)
                    {
                        Position++;
                    }
                    continue;
                }
                break;
            }
        }

        /// <summary>
        /// 跳过空白与注释后窥探下一个字节，游标不前进。
        /// Peeks the next byte after skipping whitespace and comments without advancing the cursor.
        /// </summary>
        /// <returns>下一个字节值；已到末尾返回 -1。The next byte value; -1 when the end is reached.</returns>
        public int PeekByte()
        {
            SkipWhitespaceAndComments();
            return Position < _data.Length ? _data[Position] : -1;
        }

        /// <summary>
        /// 窥探下一个关键字并在读取后把游标还原到调用前的位置。
        /// Peeks the next keyword and then restores the cursor to its position before the call.
        /// </summary>
        /// <returns>关键字文本；无关键字可读时为空字符串。The keyword text; an empty string when no keyword can be read.</returns>
        public string PeekKeyword()
        {
            int saved = Position;
            SkipWhitespaceAndComments();
            int start = Position;
            string keyword = ReadKeywordRaw();
            Position = saved;
            return keyword;
        }

        /// <summary>
        /// 跳过空白与注释后读取一个关键字并前进游标。
        /// Reads one keyword after skipping whitespace and comments, advancing the cursor.
        /// </summary>
        /// <returns>关键字文本。The keyword text.</returns>
        public string ReadKeyword()
        {
            SkipWhitespaceAndComments();
            return ReadKeywordRaw();
        }

        /// <summary>
        /// 从当前位置读取到下一个定界符为止的 ASCII 词元；若首字节即定界符则跳过它并返回空串，避免死循环。
        /// Reads the ASCII token from the current position up to the next delimiter; if the first byte is already a delimiter it is skipped and an empty string is returned to avoid an infinite loop.
        /// </summary>
        /// <returns>词元文本。The token text.</returns>
        private string ReadKeywordRaw()
        {
            int start = Position;
            while (Position < _data.Length && !IsDelimiter(_data[Position]))
            {
                Position++;
            }
            if (Position == start)
            {
                Position++;
                return string.Empty;
            }
            return Encoding.ASCII.GetString(_data, start, Position - start);
        }

        /// <summary>
        /// 若当前位置恰好是指定关键字则消费它。
        /// Consumes the given keyword when the current position matches it exactly.
        /// </summary>
        /// <param name="keyword">期望的关键字。The expected keyword.</param>
        /// <returns>匹配并消费返回 true。True when matched and consumed.</returns>
        public bool TryReadKeyword(string keyword)
        {
            SkipWhitespaceAndComments();
            if (MatchesKeyword(Position, keyword))
            {
                Position += keyword.Length;
                return true;
            }
            return false;
        }

        /// <summary>
        /// 检查缓冲中某偏移处是否为完整关键字（其后必须紧跟定界符或到达末尾），不改变游标。
        /// Checks whether a given offset in the buffer holds a complete keyword (it must be followed by a delimiter or the end) without changing the cursor.
        /// </summary>
        /// <param name="position">待检查的字节偏移。The byte offset to check.</param>
        /// <param name="keyword">期望的关键字。The expected keyword.</param>
        /// <returns>匹配返回 true。True when matched.</returns>
        public bool MatchesKeyword(int position, string keyword)
        {
            if (position < 0 || position + keyword.Length > _data.Length) return false;
            for (int i = 0; i < keyword.Length; i++)
            {
                if (_data[position + i] != (byte)keyword[i]) return false;
            }
            int after = position + keyword.Length;
            return after >= _data.Length || IsDelimiter(_data[after]);
        }

        /// <summary>
        /// 解析一个直接对象（名称、字符串、数组、字典、布尔、null 或数字/引用）。
        /// Parses a direct object (name, string, array, dictionary, boolean, null, or number/reference).
        /// </summary>
        /// <returns>解析得到的 PDF 对象。The parsed PDF object.</returns>
        /// <exception cref="InvalidDataException">数据提前结束或遇到无法识别的词元时抛出。Thrown when the data ends prematurely or an unrecognized token is encountered.</exception>
        /// <remarks>以数字开头的词元会先尝试匹配「对象号 代号 R」的引用形式，失败则回退为普通数值并还原游标。
        /// Tokens starting with a digit are first matched against the "object-number generation R" reference form; on failure they fall back to a plain number and the cursor is restored.</remarks>
        public PdfObject ParseObject()
        {
            SkipWhitespaceAndComments();
            if (Position >= _data.Length)
            {
                throw new InvalidDataException("Unexpected end of PDF data.");
            }

            byte value = _data[Position];
            switch (value)
            {
                case (byte)'/':
                    Position++;
                    return new PdfName(ReadNameRaw());
                case (byte)'(':
                    return ReadLiteralString();
                case (byte)'<':
                    if (Position + 1 < _data.Length && _data[Position + 1] == (byte)'<')
                    {
                        Position += 2;
                        return ReadDictionary();
                    }
                    return ReadHexString();
                case (byte)'[':
                    return ReadArray();
                case (byte)'t':
                    if (MatchesKeyword(Position, "true")) { Position += 4; return PdfBoolean.True; }
                    break;
                case (byte)'f':
                    if (MatchesKeyword(Position, "false")) { Position += 5; return PdfBoolean.False; }
                    break;
                case (byte)'n':
                    if (MatchesKeyword(Position, "null")) { Position += 4; return PdfNull.Instance; }
                    break;
            }

            if (value == (byte)'+' || value == (byte)'-' || value == (byte)'.' || (value >= (byte)'0' && value <= (byte)'9'))
            {
                return ReadNumberOrReference();
            }

            throw new InvalidDataException("Unexpected PDF token at position " + Position.ToString(CultureInfo.InvariantCulture) + ".");
        }

        /// <summary>
        /// 读取名称正文并展开 #HH 转义序列（调用方已消费前导斜杠）。
        /// Reads the name body and expands #HH escape sequences (the caller has already consumed the leading slash).
        /// </summary>
        /// <returns>不含斜杠的名称正文。The name body without the slash.</returns>
        private string ReadNameRaw()
        {
            StringBuilder builder = new StringBuilder();
            while (Position < _data.Length)
            {
                byte value = _data[Position];
                if (IsDelimiter(value)) break;
                if (value == (byte)'#')
                {
                    int high;
                    int low;
                    if (Position + 2 < _data.Length && TryHexValue(_data[Position + 1], out high) && TryHexValue(_data[Position + 2], out low))
                    {
                        builder.Append((char)((high << 4) | low));
                        Position += 3;
                        continue;
                    }
                }
                builder.Append((char)value);
                Position++;
            }
            return builder.ToString();
        }

        /// <summary>
        /// 读取括号字面串，处理嵌套括号、命名转义、八进制转义以及 CR/LF 续行。
        /// Reads a parenthesized literal string, handling nested parentheses, named escapes, octal escapes, and CR/LF line continuations.
        /// </summary>
        /// <returns>解析出的字符串对象（非十六进制形式）。The parsed string object (non-hex form).</returns>
        /// <remarks>遇到括号不配对的畸形数据时会在缓冲末尾停止，而不是抛出异常。
        /// Malformed data with unbalanced parentheses stops at the end of the buffer instead of throwing.</remarks>
        private PdfString ReadLiteralString()
        {
            Position++;
            List<byte> bytes = new List<byte>();
            int depth = 1;
            while (Position < _data.Length)
            {
                byte value = _data[Position++];
                if (value == (byte)'\\')
                {
                    if (Position >= _data.Length) break;
                    byte escaped = _data[Position++];
                    switch (escaped)
                    {
                        case (byte)'n': bytes.Add(10); break;
                        case (byte)'r': bytes.Add(13); break;
                        case (byte)'t': bytes.Add(9); break;
                        case (byte)'b': bytes.Add(8); break;
                        case (byte)'f': bytes.Add(12); break;
                        case (byte)'(': bytes.Add((byte)'('); break;
                        case (byte)')': bytes.Add((byte)')'); break;
                        case (byte)'\\': bytes.Add((byte)'\\'); break;
                        case 13:
                            if (Position < _data.Length && _data[Position] == 10) Position++;
                            break;
                        case 10:
                            break;
                        default:
                            if (escaped >= (byte)'0' && escaped <= (byte)'7')
                            {
                                int octal = escaped - (byte)'0';
                                int count = 1;
                                while (count < 3 && Position < _data.Length
                                    && _data[Position] >= (byte)'0' && _data[Position] <= (byte)'7')
                                {
                                    octal = octal * 8 + (_data[Position] - (byte)'0');
                                    Position++;
                                    count++;
                                }
                                bytes.Add((byte)(octal & 0xFF));
                            }
                            else
                            {
                                bytes.Add(escaped);
                            }
                            break;
                    }
                    continue;
                }
                if (value == (byte)'(')
                {
                    depth++;
                    bytes.Add(value);
                    continue;
                }
                if (value == (byte)')')
                {
                    depth--;
                    if (depth == 0) break;
                    bytes.Add(value);
                    continue;
                }
                bytes.Add(value);
            }
            return new PdfString(bytes.ToArray(), false);
        }

        /// <summary>
        /// 读取 &lt;...&gt; 形式的十六进制串，忽略其中任意非十六进制字符（调用方已判定首个尖括号）。
        /// Reads a &lt;...&gt; hex string, ignoring any non-hex digits within it (the caller has already confirmed the first angle bracket).
        /// </summary>
        /// <returns>解析出的字符串对象（十六进制形式）。The parsed string object (hex form).</returns>
        private PdfString ReadHexString()
        {
            Position++;
            List<byte> bytes = new List<byte>();
            int high = -1;
            while (Position < _data.Length)
            {
                byte value = _data[Position++];
                if (value == (byte)'>') break;
                int digit;
                if (TryHexValue(value, out digit))
                {
                    if (high < 0) high = digit;
                    else
                    {
                        bytes.Add((byte)((high << 4) | digit));
                        high = -1;
                    }
                }
            }
            if (high >= 0) bytes.Add((byte)(high << 4));
            return new PdfString(bytes.ToArray(), true);
        }

        /// <summary>
        /// 读取数组，直到 ']' 或数据结束。
        /// Reads an array until ']' or the end of the data.
        /// </summary>
        /// <returns>解析出的数组对象（调用方已消费 '['）。The parsed array object (the caller has already consumed '[').</returns>
        private PdfArray ReadArray()
        {
            Position++;
            PdfArray array = new PdfArray();
            while (true)
            {
                SkipWhitespaceAndComments();
                if (Position >= _data.Length) break;
                if (_data[Position] == (byte)']')
                {
                    Position++;
                    break;
                }
                array.Add(ParseObject());
            }
            return array;
        }

        /// <summary>
        /// 读取字典，直到 &gt;&gt; 或数据结束（调用方已消费前导的 &lt;&lt;）。
        /// Reads a dictionary until &gt;&gt; or the end of the data (the caller has already consumed the leading &lt;&lt;).
        /// </summary>
        /// <returns>解析出的字典对象。The parsed dictionary object.</returns>
        /// <remarks>键不是名称时跳过该项，以容忍被截断或写入异常的字典。
        /// Entries whose key is not a name are skipped, tolerating truncated or malformed dictionaries.</remarks>
        private PdfDictionary ReadDictionary()
        {
            PdfDictionary dictionary = new PdfDictionary();
            while (true)
            {
                SkipWhitespaceAndComments();
                if (Position >= _data.Length) break;
                if (_data[Position] == (byte)'>')
                {
                    Position++;
                    if (Position < _data.Length && _data[Position] == (byte)'>') Position++;
                    break;
                }
                PdfObject key = ParseObject();
                PdfName name = key as PdfName;
                if (name == null) continue;
                PdfObject value = ParseObject();
                dictionary.Set(name.Value, value);
            }
            return dictionary;
        }

        /// <summary>
        /// 读取数值词元；若它是不含小数点的整数且其后紧跟「代号 R」，则改判为间接对象引用。
        /// Reads a numeric token; if it is an integer without a decimal point followed by "generation R", it is reinterpreted as an indirect object reference.
        /// </summary>
        /// <returns><see cref="PdfReference"/> 或 <see cref="PdfNumber"/>。A <see cref="PdfReference"/> or a <see cref="PdfNumber"/>.</returns>
        /// <exception cref="InvalidDataException">词元无法按数值解析时抛出。Thrown when the token cannot be parsed as a number.</exception>
        /// <remarks>引用匹配失败时会把游标回退到数值之后的位置，保证后续关键字仍能被正确读取。
        /// When the reference match fails, the cursor rewinds to just after the number so any following keyword is still read correctly.</remarks>
        private PdfObject ReadNumberOrReference()
        {
            int start = Position;
            if (_data[Position] == (byte)'+' || _data[Position] == (byte)'-') Position++;
            bool hasDot = false;
            while (Position < _data.Length)
            {
                byte value = _data[Position];
                if (value >= (byte)'0' && value <= (byte)'9')
                {
                    Position++;
                    continue;
                }
                if (value == (byte)'.' && !hasDot)
                {
                    hasDot = true;
                    Position++;
                    continue;
                }
                break;
            }
            string token = Encoding.ASCII.GetString(_data, start, Position - start);
            double number;
            if (!double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
            {
                throw new InvalidDataException("Invalid PDF number '" + token + "'.");
            }

            if (!hasDot && token.IndexOf('e') < 0 && token.IndexOf('E') < 0)
            {
                int objectNumber;
                if (int.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out objectNumber) && objectNumber >= 0)
                {
                    int afterFirst = Position;
                    SkipWhitespaceAndComments();
                    int generationStart = Position;
                    while (Position < _data.Length && _data[Position] >= (byte)'0' && _data[Position] <= (byte)'9')
                    {
                        Position++;
                    }
                    if (Position > generationStart)
                    {
                        string generationToken = Encoding.ASCII.GetString(_data, generationStart, Position - generationStart);
                        SkipWhitespaceAndComments();
                        if (Position < _data.Length && _data[Position] == (byte)'R' && MatchesKeywordStart(Position))
                        {
                            Position++;
                            int generation;
                            if (int.TryParse(generationToken, NumberStyles.Integer, CultureInfo.InvariantCulture, out generation))
                            {
                                return new PdfReference(objectNumber, generation);
                            }
                        }
                    }
                    Position = afterFirst;
                }
            }
            return new PdfNumber(number);
        }

        /// <summary>
        /// 判断单字符关键字 R 之后是否为定界符，避免把 Rx 之类的词误判为引用。
        /// Determines whether the single-character keyword R is followed by a delimiter, avoiding words like Rx being misread as a reference.
        /// </summary>
        /// <param name="position">R 所在偏移。The offset where R sits.</param>
        /// <returns>后随定界符或到达末尾返回 true。True when a delimiter or the end follows.</returns>
        private bool MatchesKeywordStart(int position)
        {
            int after = position + 1;
            return after >= _data.Length || IsDelimiter(_data[after]);
        }

        /// <summary>
        /// 尝试把字符按十六进制位解释。
        /// Attempts to interpret a character as a hex digit.
        /// </summary>
        /// <param name="value">ASCII 字节。The ASCII byte.</param>
        /// <param name="result">命中时为 0–15 的位值，未命中为 0。The digit value 0-15 on a hit, otherwise 0.</param>
        /// <returns>是十六进制字符返回 true。True when the character is a hex digit.</returns>
        private static bool TryHexValue(byte value, out int result)
        {
            if (value >= (byte)'0' && value <= (byte)'9') { result = value - (byte)'0'; return true; }
            if (value >= (byte)'a' && value <= (byte)'f') { result = value - (byte)'a' + 10; return true; }
            if (value >= (byte)'A' && value <= (byte)'F') { result = value - (byte)'A' + 10; return true; }
            result = 0;
            return false;
        }

        /// <summary>
        /// 读取交叉表用的十进制偏移量；缺失数字时返回 0，以便继续扫描后续条目。
        /// Reads a decimal offset used by the cross-reference table; returns 0 when no digits are present so scanning of later entries can continue.
        /// </summary>
        /// <returns>读到的偏移值。The offset value read.</returns>
        public long ReadOffset()
        {
            SkipWhitespaceAndComments();
            int start = Position;
            while (Position < _data.Length && _data[Position] >= (byte)'0' && _data[Position] <= (byte)'9')
            {
                Position++;
            }
            long value = 0;
            for (int i = start; i < Position; i++)
            {
                value = value * 10 + (_data[i] - (byte)'0');
            }
            return value;
        }

        /// <summary>
        /// 读取带符号的十进制整数。
        /// Reads a signed decimal integer.
        /// </summary>
        /// <returns>整数值。The integer value.</returns>
        /// <exception cref="InvalidDataException">当前位置没有数字时抛出。Thrown when no digit is present at the current position.</exception>
        public int ReadInteger()
        {
            SkipWhitespaceAndComments();
            bool negative = false;
            if (Position < _data.Length && (_data[Position] == (byte)'-' || _data[Position] == (byte)'+'))
            {
                negative = _data[Position] == (byte)'-';
                Position++;
            }
            int start = Position;
            int value = 0;
            while (Position < _data.Length && _data[Position] >= (byte)'0' && _data[Position] <= (byte)'9')
            {
                value = value * 10 + (_data[Position] - (byte)'0');
                Position++;
            }
            if (Position == start) throw new InvalidDataException("Expected an integer in PDF data.");
            return negative ? -value : value;
        }

        /// <summary>
        /// 原样读取指定长度的字节并前移游标。
        /// Reads the requested number of bytes verbatim and advances the cursor.
        /// </summary>
        /// <param name="count">要读取的字节数。The number of bytes to read.</param>
        /// <returns>新分配的字节副本。A newly allocated copy of the bytes.</returns>
        /// <exception cref="InvalidDataException">长度为负或超出缓冲范围时抛出。Thrown when the length is negative or exceeds the buffer bounds.</exception>
        public byte[] ReadBytes(int count)
        {
            if (count < 0 || Position + count > _data.Length) throw new InvalidDataException("PDF stream exceeds file bounds.");
            byte[] result = new byte[count];
            Buffer.BlockCopy(_data, Position, result, 0, count);
            Position += count;
            return result;
        }

        /// <summary>
        /// 跳转到对象偏移处，解析「对象号 代号 obj」头部及其后的对象（含流）。
        /// Seeks to an object offset and parses the "object-number generation obj" header and the object (including streams) that follows.
        /// </summary>
        /// <param name="offset">对象在文件中的字节偏移，通常来自交叉表。The object's byte offset within the file, usually from the cross-reference table.</param>
        /// <returns>该间接对象承载的值；流对象返回 <see cref="PdfStream"/>（未解码）。The value carried by the indirect object; stream objects return an undecoded <see cref="PdfStream"/>.</returns>
        /// <exception cref="InvalidDataException">偏移越界或头部不符合对象语法时抛出。Thrown when the offset is out of bounds or the header does not match object syntax.</exception>
        public PdfObject ParseIndirectObjectAt(long offset)
        {
            if (offset < 0 || offset >= _data.Length)
            {
                throw new InvalidDataException("PDF object offset is outside the file.");
            }
            Position = (int)offset;
            SkipWhitespaceAndComments();
            ReadInteger();
            ReadInteger();
            string keyword = ReadKeyword();
            if (keyword != "obj")
            {
                throw new InvalidDataException("Expected 'obj' at PDF object offset " + offset.ToString(CultureInfo.InvariantCulture) + ".");
            }
            PdfObject value = ParseObject();
            PdfDictionary dictionary = value as PdfDictionary;
            if (dictionary != null)
            {
                SkipWhitespaceAndComments();
                if (MatchesKeyword(Position, "stream"))
                {
                    Position += "stream".Length;
                    if (Position < _data.Length && _data[Position] == 13)
                    {
                        Position++;
                        if (Position < _data.Length && _data[Position] == 10) Position++;
                    }
                    else if (Position < _data.Length && _data[Position] == 10)
                    {
                        Position++;
                    }
                    return ReadStream(dictionary);
                }
            }
            return value;
        }

        /// <summary>
        /// 读取流内容（调用方已消费 stream 关键字与其后的换行）。
        /// Reads the stream contents (the caller has already consumed the stream keyword and its following newline).
        /// </summary>
        /// <param name="dictionary">与该流关联的字典。The dictionary associated with the stream.</param>
        /// <returns>尚未按过滤器解码的流对象。The stream object, not yet filter-decoded.</returns>
        /// <remarks>优先信任 /Length 并顺带跳过可能的尾部换行；当 /Length 缺失、越界或与其后 endstream 不符时，
        /// 回退为直接搜索 endstream 标记，并按规范裁掉末尾的一个 CR/LF/CRLF。
        /// /Length is trusted first, also skipping a possible trailing newline; when /Length is missing, out of bounds, or inconsistent with the following endstream, it falls back to searching for the endstream marker and trims one CR/LF/CRLF at the end per the specification.</remarks>
        private PdfStream ReadStream(PdfDictionary dictionary)
        {
            int start = Position;
            int length = ResolveLength(dictionary.Get("Length"));
            if (length >= 0 && start + length <= _data.Length)
            {
                int end = start + length;
                int saved = end;
                Position = end;
                SkipWhitespaceAndComments();
                if (MatchesKeyword(Position, "endstream"))
                {
                    Position += "endstream".Length;
                    byte[] data = new byte[length];
                    Buffer.BlockCopy(_data, start, data, 0, length);
                    return new PdfStream(dictionary, data, false);
                }
                Position = saved;
            }

            int marker = FindToken("endstream", start);
            if (marker < 0) throw new InvalidDataException("Unable to locate the end of a PDF stream.");
            int dataEnd = marker;
            if (dataEnd > start && _data[dataEnd - 1] == 10)
            {
                dataEnd--;
                if (dataEnd > start && _data[dataEnd - 1] == 13) dataEnd--;
            }
            else if (dataEnd > start && _data[dataEnd - 1] == 13)
            {
                dataEnd--;
            }
            byte[] fallback = new byte[Math.Max(0, dataEnd - start)];
            Buffer.BlockCopy(_data, start, fallback, 0, fallback.Length);
            Position = marker + "endstream".Length;
            return new PdfStream(dictionary, fallback, false);
        }

        /// <summary>
        /// 解析 /Length 的值，必要时通过委托解引用。
        /// Resolves the /Length value, dereferencing through the delegate when needed.
        /// </summary>
        /// <param name="lengthObject">/Length 原始对象。The raw /Length object.</param>
        /// <returns>可用长度；缺失、类型不符或超出 int 范围时返回 -1，交由调用方走标记回退路径。The usable length; -1 when missing, of the wrong type, or outside the int range, so the caller takes the marker-fallback path.</returns>
        private int ResolveLength(PdfObject lengthObject)
        {
            if (lengthObject is PdfReference && _resolve != null)
            {
                lengthObject = _resolve((PdfReference)lengthObject);
            }
            PdfNumber number = lengthObject as PdfNumber;
            if (number == null) return -1;
            if (number.Value < 0 || number.Value > int.MaxValue) return -1;
            return (int)number.Value;
        }

        /// <summary>
        /// 在当前缓冲中从指定位置起搜索标记文本。
        /// Searches the current buffer for a token string starting at the given position.
        /// </summary>
        /// <param name="token">要查找的 ASCII 标记。The ASCII token to find.</param>
        /// <param name="start">起始偏移，负值按 0 处理。The starting offset; negative values are treated as 0.</param>
        /// <returns>首次命中的偏移；未找到返回 -1。The offset of the first hit; -1 when not found.</returns>
        public int FindToken(string token, int start)
        {
            if (start < 0) start = 0;
            for (int i = start; i + token.Length <= _data.Length; i++)
            {
                bool match = true;
                for (int j = 0; j < token.Length; j++)
                {
                    if (_data[i + j] != (byte)token[j]) { match = false; break; }
                }
                if (match) return i;
            }
            return -1;
        }

        /// <summary>
        /// 在任意字节缓冲中搜索标记文本的静态版本，供尚未构造解析器时定位文件尾与 trailer 使用。
        /// Static variant that searches any byte buffer for a token, used to locate the file trailer before a parser is constructed.
        /// </summary>
        /// <param name="data">待搜索的字节，null 视为未找到。The bytes to search; null is treated as not found.</param>
        /// <param name="token">要查找的 ASCII 标记，null 或空视为未找到。The ASCII token to find; null or empty is treated as not found.</param>
        /// <param name="start">起始偏移，负值按 0 处理。The starting offset; negative values are treated as 0.</param>
        /// <returns>首次命中的偏移；未找到返回 -1。The offset of the first hit; -1 when not found.</returns>
        public static int FindToken(byte[] data, string token, int start)
        {
            if (data == null || token == null || token.Length == 0) return -1;
            if (start < 0) start = 0;
            for (int i = start; i + token.Length <= data.Length; i++)
            {
                bool match = true;
                for (int j = 0; j < token.Length; j++)
                {
                    if (data[i + j] != (byte)token[j]) { match = false; break; }
                }
                if (match) return i;
            }
            return -1;
        }
    }
}
