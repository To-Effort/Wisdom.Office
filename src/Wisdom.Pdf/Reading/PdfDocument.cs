using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Pdf
{
    /// <summary>读取 PDF 文件，提取文本、表格与文档元数据。Reads PDF files, extracts text, tables and document metadata.</summary>
    public sealed class PdfDocument
    {
        /// <summary>交叉表条目，记录对象号到其存储位置的映射。Cross-reference entry mapping an object number to its storage location.</summary>
        private sealed class XrefEntry
        {
            /// <summary>条目类型：1 为普通对象（偏移有效），2 为对象流内的压缩对象。Entry type: 1 = plain object at an offset, 2 = compressed object inside an object stream.</summary>
            public int Type;
            /// <summary>对象头的字节偏移。Byte offset of the object head.</summary>
            public long Offset;
            /// <summary>对象代号，解密时参与对象密钥派生。The generation number, used in per-object key derivation.</summary>
            public int Generation;
            /// <summary>类型 2 时所在对象流的对象号。For type 2, the object number of the containing object stream.</summary>
            public int ObjectStream;
            /// <summary>类型 2 时在对象流中的序号。For type 2, the index of the object within that stream.</summary>
            public int ObjectIndex;
        }

        /// <summary>一个交叉表区段：经典表或交叉表流，以及指向上一区段的链接。One cross-reference section: classic table or xref stream, plus the link to the previous section.</summary>
        private sealed class XrefSection
        {
            /// <summary>该区段的 trailer 字典（交叉表流时即流字典）。Trailer dictionary of this section (the stream dictionary for xref streams).</summary>
            public PdfDictionary Trailer;
            /// <summary>/Prev 指向的上一区段偏移，0 表示没有更多区段。Offset of the previous section from /Prev; 0 means no more sections.</summary>
            public long Previous;
            /// <summary>经典 trailer 的 /XRefStm 增量交叉表流偏移，0 表示无。Hybrid /XRefStm offset from a classic trailer; 0 means absent.</summary>
            public long XRefStreamOffset;
        }

        /// <summary>交叉表：对象号到其存储位置的映射。Cross-reference map from object number to its storage location.</summary>
        private readonly Dictionary<int, XrefEntry> _xref = new Dictionary<int, XrefEntry>();
        /// <summary>已释放（type 0 或 f）的对象号，阻止更早的 /Prev 段把它们复活。
        /// Released object numbers (type 0 or f) so an older /Prev section cannot resurrect them.</summary>
        private readonly HashSet<int> _freeNumbers = new HashSet<int>();
        /// <summary>本节读到但尚未生效的空闲号，等 /XRefStm 读完再登记。
        /// Free numbers seen in the current section, held back until its /XRefStm companion is read.</summary>
        private readonly List<int> _deferredFree = new List<int>();
        /// <summary>已加载对象的缓存，对象号到对象。Cache of loaded objects, keyed by object number.</summary>
        private readonly Dictionary<int, PdfObject> _objects = new Dictionary<int, PdfObject>();
        /// <summary>正在加载中的对象号，用于阻断循环引用导致的无限递归。Object numbers currently being loaded; breaks reference cycles.</summary>
        private readonly HashSet<int> _loading = new HashSet<int>();
        /// <summary>按文档顺序收集的页面列表。Pages collected in document order.</summary>
        private readonly List<PdfPage> _pages = new List<PdfPage>();
        /// <summary>完整的 PDF 文件字节。The complete PDF file bytes.</summary>
        private byte[] _data;
        /// <summary>当前生效的 trailer 字典。The effective trailer dictionary.</summary>
        private PdfDictionary _trailer;
        /// <summary>文档元数据。The document metadata.</summary>
        private PdfDocumentInfo _info = new PdfDocumentInfo();
        /// <summary>口令认证通过后的安全处理器，null 表示未加密。The security handler after authentication; null when unencrypted.</summary>
        private PdfStandardSecurityHandler _security;
        /// <summary>加密字典的对象号，0 表示字典为直接对象。The object number of the encryption dictionary; 0 when it is direct.</summary>
        private int _encryptNumber;

        /// <summary>文档只能通过 <see cref="Load(byte[], string)"/> 等静态入口创建。Documents are created only through the static Load entry points.</summary>
        private PdfDocument()
        {
        }

        /// <summary>文档信息字典中的元数据。Metadata from the document information dictionary.</summary>
        public PdfDocumentInfo Info { get { return _info; } }
        /// <summary>文档是否受口令保护。Whether the document is password protected.</summary>
        public bool IsEncrypted { get { return _security != null; } }
        /// <summary>加密算法名称；未加密时为 null。The encryption algorithm name, or null when unencrypted.</summary>
        public string EncryptionAlgorithm { get { return _security == null ? null : _security.AlgorithmName; } }
        /// <summary>生效权限；未加密或以所有者口令打开时为全部权限。The effective permissions; all permissions when unencrypted or opened as the owner.</summary>
        public PdfPermissions Permissions { get { return _security == null ? PdfPermissions.All : _security.Permissions; } }
        /// <summary>页面数量。Number of pages.</summary>
        public int PageCount { get { return _pages.Count; } }
        /// <summary>按文档顺序排列的页面；页码从 1 开始。Pages in document order; page numbers start at 1.</summary>
        public IReadOnlyList<PdfPage> Pages { get { return _pages; } }

        /// <summary>
        /// 从文件路径加载 PDF 文档。Load a PDF document from a file path.
        /// </summary>
        /// <param name="path">PDF 文件路径。Path of the PDF file.</param>
        /// <param name="password">用户或所有者口令；null 表示尝试空口令。The user or owner password; null tries the empty password.</param>
        /// <returns>已解析目录与页面树的文档对象。A document with catalog and page tree resolved.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> 为 null 时抛出。Thrown when <paramref name="path"/> is null.</exception>
        /// <exception cref="PdfPasswordException">文件受保护且口令不正确时抛出。Thrown when the file is protected and the password is wrong.</exception>
        public static PdfDocument Load(string path, string password = null)
        {
            if (path == null) throw new ArgumentNullException("path");
            return Load(File.ReadAllBytes(path), password);
        }

        /// <summary>
        /// 从流加载 PDF 文档；整个流会被读入内存，交叉表回溯要求可随机访问。Load a PDF document from a stream; the whole stream is buffered because xref traversal needs random access.
        /// </summary>
        /// <param name="stream">包含 PDF 数据的流，只读取不关闭。Stream holding PDF data; read but not closed.</param>
        /// <param name="password">用户或所有者口令；null 表示尝试空口令。The user or owner password; null tries the empty password.</param>
        /// <returns>已解析的文档对象。A parsed document.</returns>
        public static PdfDocument Load(Stream stream, string password = null)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            return Load(OfficePackageLimits.ReadAllBytes(stream), password);
        }

        /// <summary>
        /// 从内存字节加载 PDF 文档。Load a PDF document from in-memory bytes.
        /// </summary>
        /// <param name="data">完整的 PDF 字节。The complete PDF byte content.</param>
        /// <param name="password">用户或所有者口令；null 表示尝试空口令。The user or owner password; null tries the empty password.</param>
        /// <returns>已解析的文档对象。A parsed document.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="data"/> 为 null 时抛出。Thrown when <paramref name="data"/> is null.</exception>
        /// <exception cref="PdfPasswordException">文件受保护且口令不正确时抛出。Thrown when the file is protected and the password is wrong.</exception>
        public static PdfDocument Load(byte[] data, string password = null)
        {
            if (data == null) throw new ArgumentNullException("data");
            PdfDocument document = new PdfDocument();
            document.Initialize(data, password);
            return document;
        }

        /// <summary>
        /// 校验文件头、装载交叉表并遍历页面树。Validate the header, load the cross-reference and walk the page tree.
        /// </summary>
        /// <param name="data">PDF 字节，保存下来供后续按需取对象。PDF bytes, retained for lazy object access.</param>
        /// <param name="password">用户或所有者口令。The user or owner password.</param>
        /// <exception cref="InvalidDataException">不是 PDF、缺失或损坏目录、页面树时抛出。Thrown when the file is not a PDF or has no usable catalog or page tree.</exception>
        /// <exception cref="PdfPasswordException">文件受保护且口令不正确时抛出。Thrown when the file is protected and the password is wrong.</exception>
        /// <remarks>交叉表读取失败或拿不到 /Root 时会退回全文扫描重建，这是损坏文档的常见救援路径。
        /// Falls back to a full-text scan when the cross-reference is unusable.</remarks>
        private void Initialize(byte[] data, string password)
        {
            _data = data;
            int header = PdfParser.FindToken(data, "%PDF-", 0);
            if (header < 0) throw new InvalidDataException("The file is not a PDF document.");

            bool loaded = TryLoadXref();
            if (!loaded || _trailer == null || _trailer.Get("Root") == null)
            {
                _xref.Clear();
                _objects.Clear();
                RebuildXref();
            }

            PdfDictionary trailer = _trailer;
            if (trailer == null || trailer.Get("Root") == null)
            {
                throw new InvalidDataException("The PDF document does not contain a valid catalog.");
            }

            InitializeEncryption(trailer, password);

            PdfDictionary infoDictionary = Resolve(trailer.Get("Info")) as PdfDictionary;
            _info = PdfDocumentInfo.FromDictionary(infoDictionary, this);

            PdfDictionary catalog = Resolve(trailer.Get("Root")) as PdfDictionary;
            if (catalog == null) throw new InvalidDataException("The PDF document catalog is invalid.");
            PdfDictionary pages = Resolve(catalog.Get("Pages")) as PdfDictionary;
            if (pages == null) throw new InvalidDataException("The PDF page tree is missing.");
            LoadPageNode(pages, null, null, 0, 0, new HashSet<PdfDictionary>());
            for (int i = 0; i < _pages.Count; i++)
            {
                _pages[i].Number = i + 1;
            }
        }

        /// <summary>
        /// 读取 trailer 中的加密字典并认证口令；认证前已缓存的对象会被清空，随后按对象重新解密。
        /// Reads the encryption dictionary from the trailer and authenticates the password; objects cached
        /// before authentication are discarded and decrypted again on demand.
        /// </summary>
        /// <param name="trailer">当前 trailer。The effective trailer.</param>
        /// <param name="password">用户或所有者口令。The user or owner password.</param>
        private void InitializeEncryption(PdfDictionary trailer, string password)
        {
            PdfObject encrypt = trailer.Get("Encrypt");
            if (encrypt == null) return;

            PdfReference reference = encrypt as PdfReference;
            PdfDictionary dictionary = Resolve(encrypt) as PdfDictionary;
            if (dictionary == null) throw new InvalidDataException("The PDF encryption dictionary is invalid.");

            byte[] documentId = ReadFirstDocumentId(trailer);
            _security = PdfStandardSecurityHandler.Open(dictionary, documentId, password);
            _encryptNumber = reference == null ? 0 : reference.Number;
            _objects.Clear();
            _loading.Clear();
        }

        /// <summary>
        /// 读取 trailer 的第一个 /ID 串，缺失时返回空数组。
        /// Reads the first /ID string of the trailer, or an empty array when absent.
        /// </summary>
        /// <param name="trailer">当前 trailer。The effective trailer.</param>
        /// <returns>标识字节。The identifier bytes.</returns>
        private byte[] ReadFirstDocumentId(PdfDictionary trailer)
        {
            PdfArray id = Resolve(trailer.Get("ID")) as PdfArray;
            if (id == null || id.Count == 0) return new byte[0];
            PdfString text = Resolve(id[0]) as PdfString;
            return text == null ? new byte[0] : text.Bytes;
        }

        /// <summary>
        /// 从文件尾部定位 startxref 并沿 /Prev 链读全部交叉表区段。Locate startxref near the end of the file and follow the /Prev chain through all sections.
        /// </summary>
        /// <returns>至少成功读到一个区段且有 trailer 时返回 true。True when at least one section was read.</returns>
        /// <remarks>只扫描末尾 8 KB 并取最后一次出现，因为增量更新会在历史 trailer 中留下多个 startxref；
        /// visited 集合防止 /Prev 自引用造成死循环。Only the last occurrence in the final 8 KB is used, since incremental updates leave earlier markers.</remarks>
        private bool TryLoadXref()
        {
            int searchStart = Math.Max(0, _data.Length - 8192);
            int marker = -1;
            int scan = searchStart;
            while (true)
            {
                int found = PdfParser.FindToken(_data, "startxref", scan);
                if (found < 0) break;
                marker = found;
                scan = found + 9;
            }
            if (marker < 0) return false;

            PdfParser parser = new PdfParser(_data, marker + 9, ResolveReference);
            long offset = parser.ReadOffset();
            if (offset <= 0 || offset >= _data.Length) return false;

            HashSet<long> visited = new HashSet<long>();
            long current = offset;
            bool any = false;
            while (current > 0 && current < _data.Length && visited.Add(current))
            {
                XrefSection section = null;
                try
                {
                    section = ReadXrefSection(current);
                }
                catch (InvalidDataException)
                {
                    section = null;
                }
                catch (ArgumentOutOfRangeException)
                {
                    section = null;
                }
                catch (OverflowException)
                {
                    section = null;
                }
                if (section == null) break;
                any = true;
                if (_trailer == null && section.Trailer != null)
                {
                    _trailer = section.Trailer;
                }
                if (section.XRefStreamOffset > 0 && section.XRefStreamOffset < _data.Length)
                {
                    ReadXrefStreamAt(section.XRefStreamOffset);
                }
                FlushFreeNumbers();
                current = section.Previous;
            }
            return any;
        }

        /// <summary>
        /// 读取给定偏移处的交叉表区段，自动区分经典 xref 表与交叉表流。Read the section at an offset, detecting whether it is a classic table or an xref stream.
        /// </summary>
        /// <param name="offset">区段起始偏移。Section start offset.</param>
        /// <returns>区段信息；偏移越界、结构不符或解析异常时返回 null。Null when the offset is unusable or parsing fails.</returns>
        private XrefSection ReadXrefSection(long offset)
        {
            if (offset < 0 || offset >= _data.Length) return null;
            PdfParser parser = new PdfParser(_data, (int)offset, ResolveReference);
            parser.SkipWhitespaceAndComments();
            if (parser.MatchesKeyword(parser.Position, "xref"))
            {
                parser.Position += 4;
                return ReadClassicXref(parser);
            }

            PdfStream stream;
            try
            {
                stream = parser.ParseIndirectObjectAt(offset) as PdfStream;
            }
            catch (Exception)
            {
                return null;
            }
            if (stream == null) return null;
            ReadXrefStreamDictionary(stream);
            XrefSection section = new XrefSection();
            section.Trailer = stream.Dictionary;
            section.Previous = ReadLong(stream.Dictionary.Get("Prev"));
            return section;
        }

        /// <summary>
        /// 解析经典 xref 表的分节条目直至 trailer。Parse subsections of a classic xref table until the trailer.
        /// </summary>
        /// <param name="parser">定位到 xref 关键字之后的解析器，游标会被推进。Parser positioned after the xref keyword; its cursor advances.</param>
        /// <returns>含 trailer 与 /Prev、/XRefStm 的区段；数据提前结束或 trailer 非字典时返回 null。Null when data ends before a trailer.</returns>
        /// <remarks>条目数按剩余文件长度折算上限截断，避免畸形大 count 造成超长循环；只收录 n 类型条目，f 条目无对象号故跳过。
        /// The entry count is clamped to what the remaining bytes can hold, and only in-use entries are recorded.</remarks>
        private XrefSection ReadClassicXref(PdfParser parser)
        {
            while (true)
            {
                parser.SkipWhitespaceAndComments();
                if (parser.EndOfData) return null;
                if (parser.MatchesKeyword(parser.Position, "trailer"))
                {
                    parser.Position += "trailer".Length;
                    PdfDictionary trailer = parser.ParseObject() as PdfDictionary;
                    XrefSection section = new XrefSection();
                    section.Trailer = trailer;
                    section.Previous = trailer == null ? 0 : ReadLong(trailer.Get("Prev"));
                    section.XRefStreamOffset = trailer == null ? 0 : ReadLong(trailer.Get("XRefStm"));
                    return section;
                }

                long start = parser.ReadOffset();
                int count = parser.ReadInteger();
                int maxEntries = Math.Max(0, (_data.Length - parser.Position) / 8) + 16;
                if (maxEntries > PdfReadLimits.MaxXrefEntries) maxEntries = PdfReadLimits.MaxXrefEntries;
                if (count > maxEntries) count = maxEntries;
                for (int i = 0; i < count; i++)
                {
                    long entryOffset = parser.ReadOffset();
                    int generation = parser.ReadInteger();
                    string type = parser.ReadKeyword();
                    // 起始号是无上限 long，截断会回绕成小对象号；f 条目也要记下，
                    // 否则最后一次增量更新删除的对象会从更早的 /Prev 段复活。
                    // The start is an unbounded long, so truncation wraps into a small object number; the f
                    // entries have to be remembered too, otherwise an object deleted by the last incremental
                    // update is resurrected from an older /Prev section.
                    long numberLong = start + i;
                    if (numberLong > PdfReadLimits.MaxObjectNumber) break;
                    int number = (int)numberLong;
                    if (type == "n")
                    {
                        AddEntry(number, new XrefEntry { Type = 1, Offset = entryOffset, Generation = generation });
                    }
                    else if (type == "f")
                    {
                        _deferredFree.Add(number);
                    }
                }
            }
        }

        /// <summary>
        /// 读取 /XRefStm 指向的混合交叉表流；失败静默忽略。Read the hybrid cross-reference stream named by /XRefStm; failures are ignored.
        /// </summary>
        /// <param name="offset">交叉表流对象偏移。Offset of the xref stream object.</param>
        private void ReadXrefStreamAt(long offset)
        {
            try
            {
                PdfParser parser = new PdfParser(_data, (int)offset, ResolveReference);
                PdfStream stream = parser.ParseIndirectObjectAt(offset) as PdfStream;
                if (stream != null) ReadXrefStreamDictionary(stream);
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// 按 /W 字段宽度解码交叉表流正文并登记条目。Decode the xref stream body using the /W field widths and register each entry.
        /// </summary>
        /// <param name="stream">交叉表流对象。The xref stream object.</param>
        /// <remarks>字段宽度 0 表示该类型省略（按规范默认为类型 1）；/Index 缺省时以 /Size 生成单一区段。
        /// 宽度非法、正文解码失败或数据不足时直接返回，保留已登记的部分条目。
        /// A zero width means the field is omitted (type defaults to 1); malformed widths or truncated data stop parsing without discarding what was already read.</remarks>
        private void ReadXrefStreamDictionary(PdfStream stream)
        {
            PdfDictionary dictionary = stream.Dictionary;
            PdfArray widths = Resolve(dictionary.Get("W")) as PdfArray;
            if (widths == null || widths.Count < 3) return;

            int[] widthValues = new int[widths.Count];
            for (int i = 0; i < widths.Count; i++)
            {
                PdfNumber number = Resolve(widths[i]) as PdfNumber;
                if (number == null) return;
                widthValues[i] = (int)number.Value;
                if (widthValues[i] < 0 || widthValues[i] > 8) return;
            }

            List<long> ranges = new List<long>();
            PdfArray index = Resolve(dictionary.Get("Index")) as PdfArray;
            if (index != null)
            {
                for (int i = 0; i + 1 < index.Count; i += 2)
                {
                    PdfNumber a = Resolve(index[i]) as PdfNumber;
                    PdfNumber b = Resolve(index[i + 1]) as PdfNumber;
                    if (a == null || b == null) continue;
                    ranges.Add((long)a.Value);
                    ranges.Add((long)b.Value);
                }
            }
            else
            {
                PdfNumber size = Resolve(dictionary.Get("Size")) as PdfNumber;
                ranges.Add(0);
                ranges.Add(size == null ? 0 : (long)size.Value);
            }
            if (ranges.Count == 0) return;

            byte[] data;
            try
            {
                data = DecodeStream(stream);
            }
            catch (Exception)
            {
                return;
            }

            int rowLength = 0;
            for (int i = 0; i < widthValues.Length; i++) rowLength += widthValues[i];
            if (rowLength <= 0) return;

            int position = 0;
            for (int r = 0; r < ranges.Count; r += 2)
            {
                long start = ranges[r];
                long count = ranges[r + 1];
                for (long i = 0; i < count; i++)
                {
                    // 先减后比避免 position + rowLength 溢出；条目号越出 int 范围即收尾。
                    // Subtract before comparing to avoid overflow; stop once the entry number leaves int range.
                    if (position > data.Length - rowLength) return;
                    long numberLong = start + i;
                    if (numberLong > PdfReadLimits.MaxObjectNumber) return;
                    long[] fields = new long[widthValues.Length];
                    for (int f = 0; f < widthValues.Length; f++)
                    {
                        long value = 0;
                        for (int b = 0; b < widthValues[f]; b++)
                        {
                            value = (value << 8) | data[position++];
                        }
                        fields[f] = value;
                    }
                    long type = widthValues[0] == 0 ? 1 : fields[0];
                    int objectNumber = (int)(start + i);
                    if (type == 1)
                    {
                        AddEntry(objectNumber, new XrefEntry
                        {
                            Type = 1,
                            Offset = fields[1],
                            Generation = widthValues.Length > 2 ? (int)fields[2] : 0
                        });
                    }
                    else if (type == 2)
                    {
                        AddEntry(objectNumber, new XrefEntry
                        {
                            Type = 2,
                            ObjectStream = (int)fields[1],
                            ObjectIndex = (int)fields[2]
                        });
                    }
                    else if (type == 0)
                    {
                        // 已释放条目只攒下号码，本节读完（含 /XRefStm）后才由 AddEntry 拦住更早段的同号条目。
                        // A released entry only collects the number; once the section, /XRefStm included, has
                        // been read, AddEntry blocks the same number coming from an older section.
                        _deferredFree.Add(objectNumber);
                    }
                }
            }
        }

        /// <summary>
        /// 把本节攒下的空闲号正式登记为屏蔽。混合交叉表的经典段里那些 f 条目只是占位，
        /// 同一对象真正活着的那份在 /XRefStm 流中，所以必须等伴生流读完再屏蔽，
        /// 且已经登记成功的号码不再回头屏蔽。
        /// Publishes the free numbers collected for this section. In a hybrid file the f rows of the classic
        /// table are only placeholders and the object really lives in the /XRefStm stream, so the numbers are
        /// withheld until that companion is read, and a number already registered is never blocked again.
        /// </summary>
        private void FlushFreeNumbers()
        {
            for (int i = 0; i < _deferredFree.Count; i++)
            {
                int number = _deferredFree[i];
                if (number > 0 && number <= PdfReadLimits.MaxObjectNumber
                    && !_xref.ContainsKey(number))
                {
                    _freeNumbers.Add(number);
                }
            }
            _deferredFree.Clear();
        }

        /// <summary>
        /// 登记一个交叉表条目。Register a cross-reference entry.
        /// </summary>
        /// <param name="number">对象号，必须为正。Object number; must be positive.</param>
        /// <param name="entry">条目内容。The entry to store.</param>
        /// <remarks>已有条目优先保留：交叉表区段按从新到旧遍历，新区的覆盖关系因此自然成立。
        /// Existing entries win, since sections are visited newest first.</remarks>
        private void AddEntry(int number, XrefEntry entry)
        {
            // 对象号封顶：畸形 /Index 不能把对象号抬到 int 上界；已释放的号不再登记。
            // Object numbers are capped so a malformed /Index cannot lift them to the int ceiling, and a
            // number already marked released is never registered again.
            if (number <= 0 || number > PdfReadLimits.MaxObjectNumber) return;
            if (_freeNumbers.Contains(number)) return;
            if (_xref.ContainsKey(number)) return;
            // 上限兜底：畸形交叉表可声明海量条目，字典无节制增长会直接吃满内存。
            // Ceiling: a malformed cross-reference table can declare millions of entries and grow this
            // dictionary until the process runs out of memory.
            if (_xref.Count >= PdfReadLimits.MaxXrefEntries) return;
            _xref[number] = entry;
        }

        /// <summary>
        /// 交叉表不可用时的救援路径：全文扫描对象头重建索引并寻找目录。Recovery path when the cross-reference is unusable: scan every object header and locate the catalog.
        /// </summary>
        /// <remarks>扫描得到的是每个对象号首次出现的偏移，重复对象（增量更新产生的旧副本）因此被新版本覆盖。
        /// Object streams cannot be scanned this way, so type 2 entries are unavailable in rebuilt documents.</remarks>
        private void RebuildXref()
        {
            Dictionary<int, XrefEntry> found = ScanIndirectObjects();
            List<int> keys = new List<int>(found.Keys);
            keys.Sort();
            for (int i = 0; i < keys.Count; i++)
            {
                // 全文扫描到的对象头同样是声明值，守同一个封顶。
                // Scanned object headers are declared values too and face the same cap.
                if (keys[i] > PdfReadLimits.MaxObjectNumber) continue;
                _xref[keys[i]] = found[keys[i]];
            }

            PdfDictionary trailer = FindLastTrailer();
            if (trailer == null) trailer = new PdfDictionary();

            if (trailer.Get("Root") == null)
            {
                for (int i = 0; i < keys.Count; i++)
                {
                    PdfDictionary dictionary = LoadObject(keys[i]) as PdfDictionary;
                    if (dictionary == null) continue;
                    PdfName type = Resolve(dictionary.Get("Type")) as PdfName;
                    if (type != null && type.Value == "Catalog")
                    {
                        trailer.Set("Root", new PdfReference(keys[i], 0));
                        break;
                    }
                }
            }
            _trailer = trailer;
        }

        /// <summary>
        /// 逐字节扫描 "n g obj" 形式的对象头，收集对象号与其起始偏移。Scan the file for "n g obj" headers and collect each object number with its start offset.
        /// </summary>
        /// <returns>对象号到头部起始偏移的字典，仅含正数对象号。Dictionary of object number to header offset; only positive numbers.</returns>
        /// <remarks>要求 obj 前后为定界符/空白，且编号与代号紧邻可解析为数字，以避开字符串或流正文中的伪 "obj"。
        /// Delimiter checks around obj reduce false hits inside string or stream content.</remarks>
        private Dictionary<int, XrefEntry> ScanIndirectObjects()
        {
            Dictionary<int, XrefEntry> result = new Dictionary<int, XrefEntry>();
            for (int i = 1; i + 3 < _data.Length; i++)
            {
                if (_data[i] != (byte)'o' || _data[i + 1] != (byte)'b' || _data[i + 2] != (byte)'j') continue;
                int after = i + 3;
                if (after < _data.Length && !PdfParser.IsDelimiter(_data[after])) continue;
                if (!PdfParser.IsWhitespace(_data[i - 1])) continue;

                int position = i - 1;
                while (position >= 0 && PdfParser.IsWhitespace(_data[position])) position--;
                int generationEnd = position + 1;
                while (position >= 0 && _data[position] >= (byte)'0' && _data[position] <= (byte)'9') position--;
                int generationStart = position + 1;
                if (generationStart == generationEnd) continue;

                while (position >= 0 && PdfParser.IsWhitespace(_data[position])) position--;
                int numberEnd = position + 1;
                while (position >= 0 && _data[position] >= (byte)'0' && _data[position] <= (byte)'9') position--;
                int numberStart = position + 1;
                if (numberStart == numberEnd) continue;

                int number;
                string numberText = Encoding.ASCII.GetString(_data, numberStart, numberEnd - numberStart);
                if (!int.TryParse(numberText, NumberStyles.Integer, CultureInfo.InvariantCulture, out number)) continue;
                if (number <= 0) continue;
                int generation;
                string generationText = Encoding.ASCII.GetString(_data, generationStart, generationEnd - generationStart);
                if (!int.TryParse(generationText, NumberStyles.Integer, CultureInfo.InvariantCulture, out generation)) generation = 0;
                result[number] = new XrefEntry { Type = 1, Offset = numberStart, Generation = generation };
            }
            return result;
        }

        /// <summary>
        /// 找出文件中最后一个能解析成字典的 trailer，作为重建时的目录来源。Find the last trailer that parses as a dictionary, used as the catalog source when rebuilding.
        /// </summary>
        /// <returns>最后一次成功解析的 trailer 字典；一个都没有时返回 null。The last successfully parsed trailer, or null.</returns>
        private PdfDictionary FindLastTrailer()
        {
            PdfDictionary result = null;
            int start = 0;
            while (true)
            {
                int index = PdfParser.FindToken(_data, "trailer", start);
                if (index < 0) break;
                try
                {
                    PdfParser parser = new PdfParser(_data, index + 7, ResolveReference);
                    PdfDictionary dictionary = parser.ParseObject() as PdfDictionary;
                    if (dictionary != null) result = dictionary;
                }
                catch (Exception)
                {
                }
                start = index + 7;
            }
            return result;
        }

        /// <summary>
        /// 反复解引用直到得到非引用对象。Dereference repeatedly until the value is not a reference.
        /// </summary>
        /// <param name="value">可能是引用的对象。The possibly indirect value.</param>
        /// <returns>最终对象；引用指向缺失对象或超过 64 层时返回当前值。The resolved value, or the current one when an indirection is missing or exceeds the 64-step guard.</returns>
        internal PdfObject Resolve(PdfObject value)
        {
            int guard = 0;
            while (value is PdfReference && guard++ < 64)
            {
                value = LoadObject(((PdfReference)value).Number);
            }
            return value;
        }

        /// <summary>
        /// 供解析器使用的引用回调适配。Reference callback adapter for the parser.
        /// </summary>
        /// <param name="reference">待解析的引用。The reference to resolve.</param>
        /// <returns>引用指向的对象或 null。The target object, or null.</returns>
        private PdfObject ResolveReference(PdfReference reference)
        {
            return LoadObject(reference.Number);
        }

        /// <summary>
        /// 按对象号惰性加载对象，带缓存与循环引用保护。Lazily load an object by number, with caching and cycle protection.
        /// </summary>
        /// <param name="number">交叉表中的对象号。The object number in the cross-reference.</param>
        /// <returns>对象；交叉表无此条目、正在加载中或解析失败时返回 null。Null when unknown, currently loading (a cycle) or unreadable.</returns>
        /// <remarks>失败结果不写入缓存，因此后续访问仍会重试；正在加载中的对象返回 null 而非抛异常，
        /// 使环形页面树只丢失一条边而不会崩溃。Failures are not cached, so later access may retry; a cyclic load yields null instead of an exception.</remarks>
        internal PdfObject LoadObject(int number)
        {
            PdfObject cached;
            if (_objects.TryGetValue(number, out cached)) return cached;
            XrefEntry entry;
            if (!_xref.TryGetValue(number, out entry)) return null;
            if (!_loading.Add(number)) return null;
            try
            {
                PdfObject result = null;
                if (entry.Type == 1)
                {
                    // 64 位偏移直接 (int) 截断会翻成负数或落到文件中部，先做范围校验。
                    // Casting a 64-bit offset straight to int can wrap negative or land mid-file, so the
                    // offset is range-checked first.
                    if (entry.Offset < 0 || entry.Offset > _data.Length) return null;
                    PdfParser parser = new PdfParser(_data, (int)entry.Offset, ResolveReference);
                    result = parser.ParseIndirectObjectAt(entry.Offset);
                    if (result != null && _security != null && number != _encryptNumber)
                    {
                        DecryptObjectGraph(result, number, entry.Generation);
                    }
                }
                else if (entry.Type == 2)
                {
                    result = LoadFromObjectStream(entry.ObjectStream, entry.ObjectIndex, number);
                }
                if (result != null) _objects[number] = result;
                return result;
            }
            catch (Exception)
            {
                return null;
            }
            finally
            {
                _loading.Remove(number);
            }
        }

        /// <summary>
        /// 从对象流（/ObjStm）中取出指定对象。Extract an object from an object stream.
        /// </summary>
        /// <param name="streamNumber">对象流的对象号。Object number of the stream.</param>
        /// <param name="objectIndex">交叉表记录的流内序号。Index recorded in the cross-reference.</param>
        /// <param name="targetNumber">目标对象号，用于校验。Target object number, used for verification.</param>
        /// <returns>流内对象；流缺失、正文为空或定位不到时返回 null。Null when the stream is missing, empty or the object cannot be located.</returns>
        /// <remarks>先按对象号在 "号 偏移" 表头中精确匹配，匹配不到才回退用 /ObjStm 序号，以兼容只写对序号的生成器；
        /// 条目数按正文长度上限截断，防止畸形 /N 造成越界读取。The number is matched first and the index only used as a fallback for writers that record only the index.</remarks>
        private PdfObject LoadFromObjectStream(int streamNumber, int objectIndex, int targetNumber)
        {
            PdfStream stream = LoadObject(streamNumber) as PdfStream;
            if (stream == null) return null;
            PdfDictionary dictionary = stream.Dictionary;
            int count = GetInt(dictionary.Get("N"), 0);
            int first = GetInt(dictionary.Get("First"), 0);
            byte[] data = DecodeStream(stream);
            if (data.Length == 0) return null;
            int maxObjects = Math.Min(PdfReadLimits.MaxObjectStreamObjects, data.Length / 4 + 1);
            if (count > maxObjects) count = maxObjects;

            PdfParser header = new PdfParser(data, 0, ResolveReference);
            int offset = -1;
            for (int i = 0; i < count; i++)
            {
                int number = header.ReadInteger();
                int relative = header.ReadInteger();
                if (number == targetNumber) offset = relative;
            }
            if (offset < 0 && objectIndex >= 0)
            {
                header.Position = 0;
                for (int i = 0; i <= objectIndex && i < count; i++)
                {
                    header.ReadInteger();
                    offset = header.ReadInteger();
                }
            }
            if (offset < 0) return null;
            // /First 与相对偏移都来自声明值，先减后比：相加回绕出的小正数会绕过范围检查，
            // 于是从流中间随便找个位置解析，返回的是毫不相干的对象。
            // Both /First and the relative offset come from declared values, so bounds are checked by
            // subtraction: a wrapped sum that looks small slips past the range test and the parser then
            // hands back an unrelated object decoded from the middle of the stream.
            if (first < 0 || offset > data.Length - first) return null;
            int position = first + offset;
            if (position >= data.Length) return null;
            PdfParser parser = new PdfParser(data, position, ResolveReference);
            return parser.ParseObject();
        }

        /// <summary>
        /// 递归解密对象图中的字符串与流数据；引用不在此展开，由各自的加载过程解密。
        /// Recursively decrypts strings and stream data in an object graph; references are not expanded here
        /// because each referenced object is decrypted while being loaded.
        /// </summary>
        /// <param name="value">待解密对象。The object to decrypt.</param>
        /// <param name="number">所属对象号。The owning object number.</param>
        /// <param name="generation">所属代号。The owning generation.</param>
        private void DecryptObjectGraph(PdfObject value, int number, int generation)
        {
            PdfString text = value as PdfString;
            if (text != null)
            {
                text.ReplaceBytes(_security.DecryptString(text.Bytes, number, generation));
                return;
            }
            PdfArray array = value as PdfArray;
            if (array != null)
            {
                for (int i = 0; i < array.Count; i++) DecryptObjectGraph(array[i], number, generation);
                return;
            }
            PdfStream stream = value as PdfStream;
            if (stream != null)
            {
                foreach (KeyValuePair<string, PdfObject> pair in stream.Dictionary)
                {
                    DecryptObjectGraph(pair.Value, number, generation);
                }
                bool metadata = PdfStandardSecurityHandler.IsMetadataStream(stream.Dictionary);
                if (!(metadata && !_security.EncryptMetadata))
                {
                    stream.ReplaceData(_security.DecryptStream(stream.Data, number, generation));
                }
                return;
            }
            PdfDictionary dictionary = value as PdfDictionary;
            if (dictionary != null)
            {
                foreach (KeyValuePair<string, PdfObject> pair in dictionary)
                {
                    DecryptObjectGraph(pair.Value, number, generation);
                }
            }
        }

        /// <summary>
        /// 按过滤器链解码流内容。Decode a stream body through its filter chain.
        /// </summary>
        /// <param name="stream">待解码的流。The stream to decode.</param>
        /// <returns>解码后的字节。The decoded bytes.</returns>
        internal byte[] DecodeStream(PdfStream stream)
        {
            return PdfStreamDecoder.Decode(stream, Resolve);
        }

        /// <summary>
        /// 把对象按整数读取。Read an object as an integer.
        /// </summary>
        /// <param name="value">可能是引用的对象。The possibly indirect value.</param>
        /// <param name="fallback">非数值时返回的默认值。Fallback when the value is not a number.</param>
        /// <returns>取整后的数值或默认值。The truncated number, or the fallback.</returns>
        private int GetInt(PdfObject value, int fallback)
        {
            PdfNumber number = Resolve(value) as PdfNumber;
            return number == null ? fallback : (int)number.Value;
        }

        /// <summary>
        /// 把对象按长整数读取，用于交叉表偏移。Read an object as a long, used for cross-reference offsets.
        /// </summary>
        /// <param name="value">可能是引用的对象。The possibly indirect value.</param>
        /// <returns>数值；非数值时返回 0。The number, or 0 when absent.</returns>
        private long ReadLong(PdfObject value)
        {
            PdfNumber number = Resolve(value) as PdfNumber;
            return number == null ? 0 : (long)number.Value;
        }

        /// <summary>
        /// 递归遍历页面树，把 /Resources、/MediaBox、/Rotate 沿继承链下传。Walk the page tree recursively, passing inherited resources, media box and rotation down the chain.
        /// </summary>
        /// <param name="node">当前节点字典。The current node dictionary.</param>
        /// <param name="inheritedResources">父级资源字典。Resources inherited from the parent.</param>
        /// <param name="inheritedMediaBox">父级媒体框（可能仍是引用）。Media box inherited from the parent, possibly a reference.</param>
        /// <param name="inheritedRotation">父级旋转角度。Rotation inherited from the parent.</param>
        /// <param name="depth">当前深度，超过 <see cref="PdfReadLimits.MaxPageTreeDepth"/> 即停止。Current depth; recursion stops at the configured limit.</param>
        /// <param name="visited">已访问节点集合，阻断互相指向的页面树。Visited nodes; blocks a page tree that points back at itself.</param>
        /// <remarks>有 /Kids 或 /Type 为 Pages 即视为中间节点，否则当作页面收集；因此缺少 /Type 的手写文档也能读出页面。
        /// A node is treated as an intermediate page-tree node whenever it has /Kids, so documents lacking /Type still parse.</remarks>
        private void LoadPageNode(PdfDictionary node, PdfDictionary inheritedResources, PdfObject inheritedMediaBox, int inheritedRotation,
            int depth, HashSet<PdfDictionary> visited)
        {
            if (node == null || depth > PdfReadLimits.MaxPageTreeDepth || !visited.Add(node)) return;

            PdfDictionary resources = Resolve(node.Get("Resources")) as PdfDictionary;
            if (resources == null) resources = inheritedResources;
            PdfObject mediaBox = node.Get("MediaBox") ?? inheritedMediaBox;
            PdfNumber rotationNumber = Resolve(node.Get("Rotate")) as PdfNumber;
            int rotation = rotationNumber == null ? inheritedRotation : (int)rotationNumber.Value;

            PdfArray kids = Resolve(node.Get("Kids")) as PdfArray;
            bool isPages = kids != null;
            if (!isPages)
            {
                PdfName type = Resolve(node.Get("Type")) as PdfName;
                isPages = type != null && type.Value == "Pages";
            }

            if (!isPages)
            {
                _pages.Add(new PdfPage(this, node, resources, mediaBox, rotation));
                return;
            }

            for (int i = 0; i < kids.Count; i++)
            {
                PdfDictionary kid = Resolve(kids[i]) as PdfDictionary;
                if (kid != null) LoadPageNode(kid, resources, mediaBox, rotation, depth + 1, visited);
            }
        }

        /// <summary>
        /// 提取全文文本，页与页之间以换行分隔。Extract all text, one page per block separated by newlines.
        /// </summary>
        /// <returns>拼接后的全文。The concatenated text.</returns>
        public string ExtractText()
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < _pages.Count; i++)
            {
                if (i > 0) builder.Append('\n');
                builder.Append(_pages[i].ExtractText());
            }
            return builder.ToString();
        }

        /// <summary>
        /// 把文档转换为通用的 OfficeDocument：每页一个节，表格转 Table、行外文本转段落，并按纵向位置排序。
        /// Converts the document to a generic OfficeDocument: one section per page, tables become Table blocks and out-of-table lines become paragraphs, ordered by vertical position.
        /// </summary>
        /// <returns>转换得到的 OfficeDocument。The resulting OfficeDocument.</returns>
        public OfficeDocument ToOfficeDocument()
        {
            OfficeDocument document = new OfficeDocument();
            document.Metadata.Title = _info.Title ?? string.Empty;
            document.Metadata.Author = _info.Author ?? string.Empty;
            document.Metadata.Subject = _info.Subject ?? string.Empty;
            document.Metadata.Keywords = _info.Keywords ?? string.Empty;
            if (!string.IsNullOrEmpty(_info.Creator)) document.Metadata.Creator = _info.Creator;

            for (int i = 0; i < _pages.Count; i++)
            {
                PdfPage page = _pages[i];
                PageSetup setup = new PageSetup();
                setup.Width = page.Width;
                setup.Height = page.Height;
                setup.MarginLeft = 36;
                setup.MarginRight = 36;
                setup.MarginTop = 36;
                setup.MarginBottom = 36;
                Section section = document.AddSection(setup);

                IReadOnlyList<PdfTable> tables = page.GetTables();
                IReadOnlyList<PdfTextLine> lines = page.GetTextLines();

                List<object> items = new List<object>();
                for (int t = 0; t < tables.Count; t++) items.Add(tables[t]);
                for (int l = 0; l < lines.Count; l++)
                {
                    if (IsInsideTable(lines[l], tables)) continue;
                    if (string.IsNullOrEmpty(lines[l].Text)) continue;
                    items.Add(lines[l]);
                }
                items.Sort(delegate(object a, object b)
                {
                    double ya = a is PdfTable ? ((PdfTable)a).Y : ((PdfTextLine)a).Y;
                    double yb = b is PdfTable ? ((PdfTable)b).Y : ((PdfTextLine)b).Y;
                    return ya.CompareTo(yb);
                });

                for (int item = 0; item < items.Count; item++)
                {
                    PdfTable table = items[item] as PdfTable;
                    if (table != null)
                    {
                        section.Blocks.Add(ConvertTable(table));
                        continue;
                    }
                    PdfTextLine line = (PdfTextLine)items[item];
                    Paragraph paragraph = new Paragraph(line.Text);
                    paragraph.SpaceAfter = 2;
                    section.Blocks.Add(paragraph);
                }
            }
            return document;
        }

        private static bool IsInsideTable(PdfTextLine line, IReadOnlyList<PdfTable> tables)
        {
            double centerX = line.X + line.Width / 2;
            double centerY = line.Y + line.Height / 2;
            for (int i = 0; i < tables.Count; i++)
            {
                PdfTable table = tables[i];
                if (centerX >= table.X - 1 && centerX <= table.X + table.Width + 1
                    && centerY >= table.Y - 1 && centerY <= table.Y + table.Height + 1)
                {
                    return true;
                }
            }
            return false;
        }

        private static Table ConvertTable(PdfTable source)
        {
            Table table = new Table();
            double columnWidth = source.ColumnCount > 0 ? source.Width / source.ColumnCount : source.Width;
            for (int c = 0; c < source.ColumnCount; c++)
            {
                table.Columns.Add(new TableColumn(columnWidth));
            }
            for (int r = 0; r < source.RowCount; r++)
            {
                TableRow row = table.AddRow();
                row.IsHeader = false;
                for (int c = 0; c < source.ColumnCount; c++)
                {
                    TableCell cell = new TableCell();
                    cell.Blocks.Add(new Paragraph(source.Cells[r, c] ?? string.Empty));
                    row.Cells.Add(cell);
                }
            }
            return table;
        }
    }

    /// <summary>PDF 信息字典中的文档级元数据。Document level metadata stored in the PDF information dictionary.</summary>
    public sealed class PdfDocumentInfo
    {
        /// <summary>文档标题。The document title.</summary>
        public string Title { get; internal set; }
        /// <summary>文档作者。The document author.</summary>
        public string Author { get; internal set; }
        /// <summary>文档主题。The document subject.</summary>
        public string Subject { get; internal set; }
        /// <summary>关键词列表。The keyword list.</summary>
        public string Keywords { get; internal set; }
        /// <summary>创建文档的应用程序。The application that created the original document.</summary>
        public string Creator { get; internal set; }
        /// <summary>生成 PDF 的应用程序。The application that produced the PDF.</summary>
        public string Producer { get; internal set; }
        /// <summary>创建时间，未记录时为 null。The creation date, or null when absent.</summary>
        public DateTime? CreationDate { get; internal set; }
        /// <summary>修改时间，未记录时为 null。The modification date, or null when absent.</summary>
        public DateTime? ModifiedDate { get; internal set; }

        internal static PdfDocumentInfo FromDictionary(PdfDictionary dictionary, PdfDocument document)
        {
            PdfDocumentInfo info = new PdfDocumentInfo();
            if (dictionary == null) return info;
            info.Title = ReadText(dictionary.Get("Title"), document);
            info.Author = ReadText(dictionary.Get("Author"), document);
            info.Subject = ReadText(dictionary.Get("Subject"), document);
            info.Keywords = ReadText(dictionary.Get("Keywords"), document);
            info.Creator = ReadText(dictionary.Get("Creator"), document);
            info.Producer = ReadText(dictionary.Get("Producer"), document);
            info.CreationDate = ParseDate(ReadText(dictionary.Get("CreationDate"), document));
            info.ModifiedDate = ParseDate(ReadText(dictionary.Get("ModDate"), document));
            return info;
        }

        private static string ReadText(PdfObject value, PdfDocument document)
        {
            PdfString text = document.Resolve(value) as PdfString;
            if (text == null) return string.Empty;
            return DecodePdfString(text.Bytes);
        }

        internal static string DecodePdfString(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0) return string.Empty;
            if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            {
                char[] chars = new char[(bytes.Length - 2) / 2];
                for (int i = 0; i < chars.Length; i++)
                {
                    chars[i] = (char)((bytes[2 + i * 2] << 8) | bytes[3 + i * 2]);
                }
                return new string(chars);
            }
            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            {
                char[] chars = new char[(bytes.Length - 2) / 2];
                for (int i = 0; i < chars.Length; i++)
                {
                    chars[i] = (char)((bytes[2 + i * 2]) | (bytes[3 + i * 2] << 8));
                }
                return new string(chars);
            }
            char[] latin = new char[bytes.Length];
            for (int i = 0; i < bytes.Length; i++) latin[i] = (char)bytes[i];
            return new string(latin);
        }

        private static DateTime? ParseDate(string value)
        {
            if (string.IsNullOrEmpty(value)) return null;
            string text = value.Trim();
            if (text.StartsWith("D:", StringComparison.Ordinal)) text = text.Substring(2);
            if (text.Length < 4) return null;
            string digits = string.Empty;
            for (int i = 0; i < text.Length && i < 14; i++)
            {
                if (text[i] >= '0' && text[i] <= '9') digits += text[i];
                else break;
            }
            if (digits.Length < 4) return null;
            int year = ParseInt(digits, 0, 4);
            int month = Math.Max(1, Math.Min(12, digits.Length >= 6 ? ParseInt(digits, 4, 2) : 1));
            int day = digits.Length >= 8 ? ParseInt(digits, 6, 2) : 1;
            int hour = digits.Length >= 10 ? ParseInt(digits, 8, 2) : 0;
            int minute = digits.Length >= 12 ? ParseInt(digits, 10, 2) : 0;
            int second = digits.Length >= 14 ? ParseInt(digits, 12, 2) : 0;
            try
            {
                int maxDay = DateTime.DaysInMonth(year < 1 ? 1 : year, month);
                day = Math.Max(1, Math.Min(maxDay, day));
                return new DateTime(year, month, day,
                    Math.Max(0, Math.Min(23, hour)), Math.Max(0, Math.Min(59, minute)), Math.Max(0, Math.Min(59, second)));
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }
        }

        private static int ParseInt(string text, int start, int length)
        {
            int value = 0;
            for (int i = start; i < start + length && i < text.Length; i++)
            {
                if (text[i] >= '0' && text[i] <= '9') value = value * 10 + (text[i] - '0');
            }
            return value;
        }
    }
}
