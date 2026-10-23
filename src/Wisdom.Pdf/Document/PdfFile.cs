using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Pdf
{
    /// <summary>
    /// 打开带口令保护的 PDF 失败时抛出的异常。
    /// Thrown when a password-protected PDF cannot be opened with the supplied password.
    /// </summary>
    public sealed class PdfPasswordException : Exception
    {
        /// <summary>使用给定消息创建异常。Creates the exception with the given message.</summary>
        /// <param name="message">异常消息。The exception message.</param>
        public PdfPasswordException(string message)
            : base(message)
        {
        }
    }

    /// <summary>
    /// 对象级 PDF 文档：按对象号读写间接对象、编辑页面树、增量保存，是表单、签章、脱敏等能力的公共底座。
    /// Object level PDF document: reads and writes indirect objects by number, edits the page tree and
    /// saves incrementally. It is the shared foundation for forms, signatures and redaction.
    /// </summary>
    /// <remarks>
    /// 与只读的 <see cref="PdfDocument"/> 不同，本类保留原始对象编号与文件字节，因此可以只追加变更（增量更新），
    /// 从而不破坏已有数字签名；全量保存会重新序列化全部对象，适合合并、拆分与清理。
    /// Unlike the read-only <see cref="PdfDocument"/>, this class preserves object numbers and the original
    /// file bytes, so changes can be appended as an incremental update without invalidating existing
    /// signatures; a full save re-serializes every object, which suits merging, splitting and clean-up.
    /// </remarks>
    public sealed partial class PdfFile
    {
        /// <summary>交叉表条目：普通对象、对象流内对象或空闲对象。A cross-reference entry: a plain object, an object inside an object stream, or a free entry.</summary>
        internal sealed class XrefEntry
        {
            /// <summary>0 空闲、1 普通对象、2 对象流内对象。0 free, 1 plain object, 2 object inside an object stream.</summary>
            public int Type;
            /// <summary>对象头的字节偏移。Byte offset of the object header.</summary>
            public long Offset;
            /// <summary>代号。The generation number.</summary>
            public int Generation;
            /// <summary>类型 2 时所在对象流的对象号。For type 2, the object stream number.</summary>
            public int ObjectStream;
            /// <summary>类型 2 时在对象流中的序号。For type 2, the index inside the object stream.</summary>
            public int ObjectIndex;
        }

        /// <summary>页面树节点快照，保留页面字典、引用与继承属性。A snapshot of a page tree node, keeping the page dictionary, reference and inherited attributes.</summary>
        internal sealed class PageNode
        {
            /// <summary>页面字典。The page dictionary.</summary>
            public PdfDictionary Dictionary;
            /// <summary>页面引用；页面字典为直接对象时可能为 null。The page reference; null when the page dictionary is a direct object.</summary>
            public PdfReference Reference;
            /// <summary>沿继承链合并后的资源字典。The resource dictionary merged along the inheritance chain.</summary>
            public PdfDictionary Resources;
            /// <summary>沿继承链得到的媒体框。The media box inherited along the chain.</summary>
            public PdfObject MediaBox;
            /// <summary>规整后的旋转角度。The normalized rotation angle.</summary>
            public int Rotation;
        }

        /// <summary>交叉表：对象号到存储位置。Cross-reference table mapping object numbers to storage locations.</summary>
        private readonly Dictionary<int, XrefEntry> _xref = new Dictionary<int, XrefEntry>();
        /// <summary>已加载对象缓存。Cache of loaded objects.</summary>
        private readonly Dictionary<int, PdfObject> _cache = new Dictionary<int, PdfObject>();
        /// <summary>正在加载的对象号，用于阻断引用环。Object numbers currently loading, breaking reference cycles.</summary>
        private readonly HashSet<int> _loading = new HashSet<int>();
        /// <summary>待写出的变更对象。Objects changed in memory and waiting to be written.</summary>
        private readonly Dictionary<int, PdfObject> _changes = new Dictionary<int, PdfObject>();
        /// <summary>已删除的对象号。Object numbers removed from the document.</summary>
        private readonly HashSet<int> _removed = new HashSet<int>();

        /// <summary>完整文件字节，新建文档时为 null。The complete file bytes; null for a newly created document.</summary>
        private byte[] _data;
        /// <summary>当前 trailer 字典。The current trailer dictionary.</summary>
        private PdfDictionary _trailer;
        /// <summary>最近一次交叉表区段的偏移。Offset of the latest cross-reference section.</summary>
        private long _lastXrefOffset;
        /// <summary>已分配的最大对象号。The largest object number allocated so far.</summary>
        private int _maxObjectNumber;
        /// <summary>是否为新文档。Whether this is a new document.</summary>
        private bool _isNew;
        /// <summary>目录对象号；直接目录会先被提升为间接对象。The catalog object number; a direct catalog is promoted to an indirect object.</summary>
        private int _catalogNumber = -1;
        /// <summary>缓存的目录字典。The cached catalog dictionary.</summary>
        private PdfDictionary _catalog;
        /// <summary>页面节点缓存，页面操作后失效。Page node cache, invalidated after page edits.</summary>
        private List<PageNode> _pageNodes;
        /// <summary>文档信息缓存。Cached document information.</summary>
        private PdfDocumentInfo _info;
        /// <summary>当前生效的加密处理器，用于写出与公开信息。The effective security handler, used for writing and for the public properties.</summary>
        private PdfStandardSecurityHandler _security;
        /// <summary>解密源文件对象所用的处理器；更换加密后仍指向旧处理器，直到保存完成。The handler used to decrypt source objects; after an encryption change it keeps pointing at the old handler until a save completes.</summary>
        private PdfStandardSecurityHandler _sourceSecurity;
        /// <summary>源文件加密字典的对象号，该对象的字符串与流不参与解密。The source encryption dictionary object number; its strings and streams are never decrypted.</summary>
        private int _sourceEncryptNumber;
        /// <summary>当前 trailer 中的 /Encrypt 引用。The /Encrypt reference in the current trailer.</summary>
        private PdfReference _encryptReference;
        /// <summary>待应用的新加密设置。Pending new encryption settings.</summary>
        private PdfEncryptionOptions _pendingEncryption;
        /// <summary>加密设置已变化，增量保存不再安全，必须全量重写。Encryption changed, so an incremental save is unsafe and a full rewrite is required.</summary>
        private bool _rewriteRequired;
        /// <summary>文件头版本（如 1.7）。The file header version, such as 1.7.</summary>
        private string _version = "1.7";
        /// <summary>是否在读取时得到过交叉表。Whether a usable cross-reference was found while reading.</summary>
        private bool _hasXref;
        /// <summary>信息字典是否有待写回的修改。Whether the information dictionary has pending changes.</summary>
        private bool _infoDirty;

        /// <summary>通过 <see cref="Load(string,string)"/> 或 <see cref="Create"/> 创建实例。Instances are created through <see cref="Load(string,string)"/> or <see cref="Create"/>.</summary>
        private PdfFile()
        {
        }

        /// <summary>
        /// 新建一个空 PDF（A4 单页页面树 + 空目录）。
        /// Creates an empty PDF with an A4 page tree and an empty catalog.
        /// </summary>
        /// <returns>可继续导入页面或添加对象的新文档。A new document ready for page imports or object additions.</returns>
        public static PdfFile Create()
        {
            PdfFile file = new PdfFile();
            file._isNew = true;
            file._version = "1.7";
            PdfDictionary pages = new PdfDictionary();
            pages.SetName("Type", "Pages");
            pages.Set("Kids", new PdfArray());
            pages.SetNumber("Count", 0);
            int pagesNumber = file.AddObject(pages);

            PdfDictionary catalog = new PdfDictionary();
            catalog.SetName("Type", "Catalog");
            catalog.Set("Pages", new PdfReference(pagesNumber, 0));
            file._catalogNumber = file.AddObject(catalog);
            file._catalog = catalog;

            PdfDictionary info = new PdfDictionary();
            info.SetString("Producer", "Wisdom.Office PDF");
            int infoNumber = file.AddObject(info);
            file._trailer = new PdfDictionary();
            file._trailer.SetNumber("Size", file._maxObjectNumber + 1);
            file._trailer.Set("Root", new PdfReference(file._catalogNumber, 0));
            file._trailer.Set("Info", new PdfReference(infoNumber, 0));
            return file;
        }

        /// <summary>
        /// 从文件加载 PDF；口令为 null 或空时尝试空口令。
        /// Loads a PDF from a file; a null or empty password tries the empty user password.
        /// </summary>
        /// <param name="path">文件路径。The file path.</param>
        /// <param name="password">用户口令。The user password.</param>
        /// <returns>已加载的文档。The loaded document.</returns>
        public static PdfFile Load(string path, string password = null)
        {
            if (path == null) throw new ArgumentNullException("path");
            return Load(File.ReadAllBytes(path), password);
        }

        /// <summary>从流加载 PDF，流内容会被完整读入内存。Loads a PDF from a stream; the content is fully buffered in memory.</summary>
        /// <param name="stream">PDF 数据流，读取后不关闭。The PDF stream; it is read but not closed.</param>
        /// <param name="password">用户口令。The user password.</param>
        /// <returns>已加载的文档。The loaded document.</returns>
        public static PdfFile Load(Stream stream, string password = null)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            return Load(OfficePackageLimits.ReadAllBytes(stream), password);
        }

        /// <summary>从字节加载 PDF。Loads a PDF from bytes.</summary>
        /// <param name="data">完整 PDF 字节。The complete PDF bytes.</param>
        /// <param name="password">用户口令。The user password.</param>
        /// <returns>已加载的文档。The loaded document.</returns>
        public static PdfFile Load(byte[] data, string password = null)
        {
            if (data == null) throw new ArgumentNullException("data");
            PdfFile file = new PdfFile();
            file.Initialize(data, password);
            return file;
        }

        /// <summary>文件头版本，例如 1.7。The file header version, for example 1.7.</summary>
        public string Version { get { return _version; } }

        /// <summary>
        /// 设置文件头版本；PDF/A-1 等合规输出需要固定版本。
        /// Sets the file header version; compliant output such as PDF/A-1 needs a fixed version.
        /// </summary>
        /// <param name="version">版本文本，如 1.4、1.7。The version text, such as 1.4 or 1.7.</param>
        public void SetVersion(string version)
        {
            if (string.IsNullOrEmpty(version)) throw new ArgumentException("Version cannot be empty.", "version");
            _version = version;
        }

        /// <summary>是否为新建且尚未保存的文档。Whether this is a new document that has not been saved yet.</summary>
        public bool IsNew { get { return _isNew; } }

        /// <summary>是否带加密字典。Whether the document has an encryption dictionary.</summary>
        public bool IsEncrypted { get { return _security != null; } }

        /// <summary>当前加密算法名称；未加密时为 null。The current encryption algorithm name; null when not encrypted.</summary>
        public string EncryptionAlgorithm { get { return _security == null ? null : _security.AlgorithmName; } }

        /// <summary>当前权限位；未加密时为全部允许。The permission flags; all allowed when not encrypted.</summary>
        public PdfPermissions Permissions { get { return _security == null ? PdfPermissions.All : _security.Permissions; } }

        /// <summary>trailer 字典。The trailer dictionary.</summary>
        public PdfDictionary Trailer { get { return _trailer; } }

        /// <summary>目录字典（/Root）。The catalog dictionary (/Root).</summary>
        public PdfDictionary Catalog
        {
            get
            {
                if (_catalog == null)
                {
                    _catalog = Resolve(_trailer == null ? null : _trailer.Get("Root")) as PdfDictionary;
                    if (_catalog == null) throw new InvalidDataException("The PDF document does not contain a valid catalog.");
                }
                return _catalog;
            }
        }

        /// <summary>目录的对象号。The object number of the catalog.</summary>
        public int CatalogNumber { get { return _catalogNumber; } }

        /// <summary>全部已登记对象号（不含已删除对象），升序排列。All registered object numbers excluding removed ones, in ascending order.</summary>
        public IEnumerable<int> ObjectNumbers
        {
            get
            {
                SortedSet<int> numbers = new SortedSet<int>();
                foreach (int number in _xref.Keys)
                {
                    if (!_removed.Contains(number)) numbers.Add(number);
                }
                foreach (int number in _changes.Keys)
                {
                    if (!_removed.Contains(number)) numbers.Add(number);
                }
                foreach (int number in _cache.Keys)
                {
                    if (!_removed.Contains(number) && number != 0) numbers.Add(number);
                }
                return numbers;
            }
        }

        /// <summary>已分配的最大对象号。The largest object number allocated.</summary>
        public int MaxObjectNumber { get { return _maxObjectNumber; } }

        /// <summary>
        /// 读取并解密指定对象。
        /// Reads and decrypts the object with the given number.
        /// </summary>
        /// <param name="number">对象号。The object number.</param>
        /// <returns>对象；不存在或无法解析时返回 null。The object, or null when absent or unreadable.</returns>
        public PdfObject GetObject(int number)
        {
            return LoadObject(number);
        }

        /// <summary>按类型读取对象。Reads an object with the expected type.</summary>
        /// <typeparam name="T">期望的 PDF 对象类型。The expected PDF object type.</typeparam>
        /// <param name="number">对象号。The object number.</param>
        /// <returns>类型匹配的对象，否则返回 null。The object when the type matches, otherwise null.</returns>
        public T GetObject<T>(int number) where T : PdfObject
        {
            return LoadObject(number) as T;
        }

        /// <summary>
        /// 新增一个间接对象并返回对象号。
        /// Adds an indirect object and returns its object number.
        /// </summary>
        /// <param name="value">待登记的对象。The object to register.</param>
        /// <returns>新对象号。The new object number.</returns>
        public int AddObject(PdfObject value)
        {
            if (value == null) throw new ArgumentNullException("value");
            int number = NextObjectNumber();
            _changes[number] = value;
            _cache[number] = value;
            return number;
        }

        /// <summary>
        /// 新增一个间接对象并返回引用。
        /// Adds an indirect object and returns a reference to it.
        /// </summary>
        /// <param name="value">待登记的对象。The object to register.</param>
        /// <returns>指向新对象的引用。A reference to the new object.</returns>
        public PdfReference AddObjectReference(PdfObject value)
        {
            return new PdfReference(AddObject(value), 0);
        }

        /// <summary>
        /// 替换指定对象号的内容并标记为已修改。
        /// Replaces the content of an object number and marks it modified.
        /// </summary>
        /// <param name="number">对象号。The object number.</param>
        /// <param name="value">新对象。The new object.</param>
        public void SetObject(int number, PdfObject value)
        {
            if (number <= 0) throw new ArgumentOutOfRangeException("number");
            if (value == null) throw new ArgumentNullException("value");
            _removed.Remove(number);
            _changes[number] = value;
            _cache[number] = value;
            if (number > _maxObjectNumber) _maxObjectNumber = number;
            if (number == _catalogNumber) _catalog = value as PdfDictionary;
        }

        /// <summary>
        /// 标记对象为已修改，使增量保存会写出该对象。
        /// Marks an object as modified so an incremental save writes it.
        /// </summary>
        /// <param name="number">对象号。The object number.</param>
        public void MarkModified(int number)
        {
            if (number <= 0) return;
            PdfObject value;
            if (_cache.TryGetValue(number, out value))
            {
                _changes[number] = value;
                return;
            }
            if (_changes.ContainsKey(number)) return;
            value = LoadObject(number);
            if (value != null) _changes[number] = value;
        }

        /// <summary>
        /// 删除对象；增量保存会写出对应的空闲条目。
        /// Removes an object; an incremental save writes a free entry for it.
        /// </summary>
        /// <param name="number">对象号。The object number.</param>
        public void RemoveObject(int number)
        {
            if (number <= 0) return;
            _cache.Remove(number);
            _changes.Remove(number);
            _xref.Remove(number);
            _removed.Add(number);
            if (number == _catalogNumber) _catalog = null;
        }

        /// <summary>判断对象号是否存在。Determines whether the object number exists.</summary>
        /// <param name="number">对象号。The object number.</param>
        /// <returns>存在返回 true。True when the object exists.</returns>
        public bool ContainsObject(int number)
        {
            if (number <= 0 || _removed.Contains(number)) return false;
            return _cache.ContainsKey(number) || _xref.ContainsKey(number) || _changes.ContainsKey(number);
        }

        /// <summary>
        /// 解引用直到得到直接对象，最多 64 层。
        /// Dereferences until a direct object is reached, with a 64-step guard.
        /// </summary>
        /// <param name="value">可能为引用的对象。The possibly indirect object.</param>
        /// <returns>最终对象。The resolved object.</returns>
        public PdfObject Resolve(PdfObject value)
        {
            int guard = 0;
            while (value is PdfReference && guard++ < 64)
            {
                value = LoadObject(((PdfReference)value).Number);
            }
            return value;
        }

        /// <summary>
        /// 解码流内容（调用时流数据已解密）。
        /// Decodes the stream body; stream data has already been decrypted by the time this runs.
        /// </summary>
        /// <param name="stream">待解码的流。The stream to decode.</param>
        /// <returns>解码后的字节。The decoded bytes.</returns>
        public byte[] DecodeStream(PdfStream stream)
        {
            return PdfStreamDecoder.Decode(stream, Resolve);
        }

        /// <summary>文档信息字典中的元数据；设置后调用 <see cref="Save(string)"/> 写回。Metadata from the document information dictionary; assign and save to persist.</summary>
        public PdfDocumentInfo Info
        {
            get
            {
                if (_info == null) _info = ReadInfo();
                return _info;
            }
            set { _info = value; _infoDirty = true; }
        }

        /// <summary>
        /// 取得或创建文档信息字典的引用。
        /// Gets or creates the reference of the document information dictionary.
        /// </summary>
        /// <returns>信息字典引用。The information dictionary reference.</returns>
        public PdfReference GetOrCreateInfoReference()
        {
            PdfObject existing = _trailer == null ? null : _trailer.Get("Info");
            PdfReference reference = existing as PdfReference;
            if (reference != null && ContainsObject(reference.Number)) return reference;
            PdfDictionary dictionary = Resolve(existing) as PdfDictionary;
            if (dictionary == null) dictionary = new PdfDictionary();
            int number = AddObject(dictionary);
            reference = new PdfReference(number, 0);
            if (_trailer == null) _trailer = new PdfDictionary();
            _trailer.Set("Info", reference);
            return reference;
        }

        /// <summary>
        /// 把 <see cref="Info"/> 写回信息字典。
        /// Writes <see cref="Info"/> back into the information dictionary.
        /// </summary>
        public void UpdateInfo()
        {
            if (_info == null) return;
            PdfDictionary info = Resolve(GetOrCreateInfoReference()) as PdfDictionary;
            if (info == null) return;
            SetOrRemove(info, "Title", _info.Title);
            SetOrRemove(info, "Author", _info.Author);
            SetOrRemove(info, "Subject", _info.Subject);
            SetOrRemove(info, "Keywords", _info.Keywords);
            SetOrRemove(info, "Creator", _info.Creator);
            SetOrRemove(info, "Producer", _info.Producer);
            if (_info.CreationDate.HasValue) info.Set("CreationDate", PdfString.Literal(PdfObjectWriter.FormatPdfDate(_info.CreationDate.Value)));
            if (_info.ModifiedDate.HasValue) info.Set("ModDate", PdfString.Literal(PdfObjectWriter.FormatPdfDate(_info.ModifiedDate.Value)));
            MarkModified(GetOrCreateInfoReference().Number);
            _infoDirty = false;
        }

        /// <summary>
        /// 删除文档加密，随后全量保存输出未加密文件。
        /// Removes encryption; a subsequent full save writes an unencrypted file.
        /// </summary>
        public void RemoveEncryption()
        {
            _pendingEncryption = null;
            _security = null;
            _encryptReference = null;
            if (_sourceEncryptNumber > 0) RemoveObject(_sourceEncryptNumber);
            if (_trailer != null) _trailer.Remove("Encrypt");
            _rewriteRequired = _sourceSecurity != null;
        }

        /// <summary>
        /// 启用密码保护；因为需要重建加密字典，此设置只对全量保存生效。
        /// Enables password protection; because the encryption dictionary is rebuilt, this applies to a full save.
        /// </summary>
        /// <param name="options">加密选项。The encryption options.</param>
        public void SetEncryption(PdfEncryptionOptions options)
        {
            if (options == null) throw new ArgumentNullException("options");
            _pendingEncryption = options;
            _security = null;
            _encryptReference = null;
            _rewriteRequired = true;
        }

        /// <summary>
        /// 以文件头版本与重新序列化的全部对象保存文件。
        /// Saves by re-serializing every object with the file header version.
        /// </summary>
        /// <param name="path">目标路径。The target path.</param>
        public void Save(string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            File.WriteAllBytes(path, SaveToBytes());
        }

        /// <summary>把全量重写的结果写入流。Writes the full rewrite into a stream.</summary>
        /// <param name="output">目标流。The target stream.</param>
        public void Save(Stream output)
        {
            if (output == null) throw new ArgumentNullException("output");
            byte[] bytes = SaveToBytes();
            output.Write(bytes, 0, bytes.Length);
        }

        /// <summary>
        /// 全量保存并返回完整文件字节。
        /// Performs a full save and returns the complete file bytes.
        /// </summary>
        /// <returns>完整 PDF 字节。The complete PDF bytes.</returns>
        public byte[] SaveToBytes()
        {
            return WriteFile(false);
        }

        /// <summary>
        /// 增量保存：只追加变更对象与新交叉表区段，保留原始字节和既有签名。
        /// Incremental save: appends only changed objects and a new cross-reference section, keeping the
        /// original bytes and existing signatures intact.
        /// </summary>
        /// <returns>追加更新后的完整文件字节。The complete file bytes after the update.</returns>
        public byte[] SaveIncrementalToBytes()
        {
            if (_isNew || _data == null || _pendingEncryption != null || _rewriteRequired) return WriteFile(false);
            return WriteFile(true);
        }

        /// <summary>把增量更新的结果写入文件。Writes an incremental update to a file.</summary>
        /// <param name="path">目标路径。The target path.</param>
        public void SaveIncremental(string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            File.WriteAllBytes(path, SaveIncrementalToBytes());
        }

        /// <summary>
        /// 初始化：定位文件头、装载交叉表、建立加密处理器并解析目录与页面树。
        /// Initializes by locating the header, loading the cross-reference, creating the security handler
        /// and resolving the catalog and page tree.
        /// </summary>
        /// <param name="data">完整 PDF 字节。The complete PDF bytes.</param>
        /// <param name="password">用户口令。The user password.</param>
        private void Initialize(byte[] data, string password)
        {
            _data = data;
            _isNew = false;

            int header = PdfParser.FindToken(data, "%PDF-", 0);
            if (header < 0) throw new InvalidDataException("The file is not a PDF document.");
            int versionEnd = header + 5;
            while (versionEnd < data.Length && versionEnd < header + 12
                && ((data[versionEnd] >= (byte)'0' && data[versionEnd] <= (byte)'9') || data[versionEnd] == (byte)'.'))
            {
                versionEnd++;
            }
            string version = Encoding.ASCII.GetString(data, header + 5, Math.Max(0, versionEnd - header - 5));
            if (!string.IsNullOrEmpty(version)) _version = version;

            _hasXref = TryLoadXref();
            if (!_hasXref || _trailer == null || _trailer.Get("Root") == null)
            {
                _xref.Clear();
                _cache.Clear();
                RebuildXref();
            }
            if (_trailer == null || _trailer.Get("Root") == null)
            {
                throw new InvalidDataException("The PDF document does not contain a valid catalog.");
            }

            InitializeEncryption(password);

            PdfObject root = _trailer.Get("Root");
            PdfReference rootReference = root as PdfReference;
            PdfDictionary catalog = Resolve(root) as PdfDictionary;
            if (catalog == null) throw new InvalidDataException("The PDF document catalog is invalid.");
            if (rootReference == null)
            {
                _catalogNumber = AddObject(catalog);
                _trailer.Set("Root", new PdfReference(_catalogNumber, 0));
            }
            else
            {
                _catalogNumber = rootReference.Number;
            }
            _catalog = catalog;

            PdfObject infoObject = _trailer.Get("Info");
            if (infoObject != null) _info = ReadInfo();
            BuildPageNodes();
        }

        /// <summary>
        /// 读取 trailer 中的加密字典并校验口令，成功后清空预解密缓存。
        /// Reads the encryption dictionary from the trailer and validates the password, clearing any
        /// pre-decryption cache afterwards.
        /// </summary>
        /// <param name="password">用户口令。The user password.</param>
        private void InitializeEncryption(string password)
        {
            PdfObject encryptObject = _trailer.Get("Encrypt");
            if (encryptObject == null) return;

            PdfReference reference = encryptObject as PdfReference;
            PdfDictionary dictionary = Resolve(encryptObject) as PdfDictionary;
            if (dictionary == null) throw new InvalidDataException("The PDF encryption dictionary is invalid.");

            byte[] documentId = ReadFirstDocumentId();
            _security = PdfStandardSecurityHandler.Open(dictionary, documentId, password);
            _sourceSecurity = _security;
            _encryptReference = reference;
            if (_encryptReference == null)
            {
                _encryptReference = new PdfReference(AddObject(dictionary), 0);
                _trailer.Set("Encrypt", _encryptReference);
            }
            _sourceEncryptNumber = _encryptReference.Number;
            _cache.Clear();
            foreach (KeyValuePair<int, PdfObject> pair in _changes) _cache[pair.Key] = pair.Value;
            _loading.Clear();
        }

        /// <summary>
        /// 从 trailer 的 /ID 数组读取第一个标识串；缺失时用空数组。
        /// Reads the first identifier string from the trailer's /ID array, or an empty array when absent.
        /// </summary>
        /// <returns>标识字节。The identifier bytes.</returns>
        private byte[] ReadFirstDocumentId()
        {
            PdfObject value = Resolve(_trailer.Get("ID"));
            PdfArray array = value as PdfArray;
            if (array == null || array.Count == 0) return new byte[0];
            PdfString text = Resolve(array[0]) as PdfString;
            return text == null ? new byte[0] : text.Bytes;
        }

        /// <summary>
        /// 沿 startxref 与 /Prev 链装载全部交叉表区段。
        /// Loads every cross-reference section along the startxref and /Prev chain.
        /// </summary>
        /// <returns>至少读到一个有效区段时返回 true。True when at least one valid section was read.</returns>
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
            _lastXrefOffset = offset;

            HashSet<long> visited = new HashSet<long>();
            long current = offset;
            bool any = false;
            while (current > 0 && current < _data.Length && visited.Add(current))
            {
                Section section = null;
                try
                {
                    section = ReadXrefSection(current);
                }
                catch (Exception)
                {
                    section = null;
                }
                if (section == null) break;
                any = true;
                if (_trailer == null && section.Trailer != null) _trailer = section.Trailer;
                if (section.XRefStreamOffset > 0 && section.XRefStreamOffset < _data.Length)
                {
                    ReadXrefStreamAt(section.XRefStreamOffset);
                }
                current = section.Previous;
            }
            return any;
        }

        /// <summary>交叉表区段。A cross-reference section.</summary>
        private sealed class Section
        {
            /// <summary>区段 trailer。The section trailer.</summary>
            public PdfDictionary Trailer;
            /// <summary>/Prev 偏移。The /Prev offset.</summary>
            public long Previous;
            /// <summary>/XRefStm 偏移。The /XRefStm offset.</summary>
            public long XRefStreamOffset;
        }

        /// <summary>
        /// 读取一个交叉表区段，自动区分经典表与交叉表流。
        /// Reads one cross-reference section, detecting classic tables and xref streams.
        /// </summary>
        /// <param name="offset">区段偏移。The section offset.</param>
        /// <returns>区段；结构不符时返回 null。The section, or null when malformed.</returns>
        private Section ReadXrefSection(long offset)
        {
            if (offset < 0 || offset >= _data.Length) return null;
            PdfParser parser = new PdfParser(_data, (int)offset, ResolveReference);
            parser.SkipWhitespaceAndComments();
            if (parser.MatchesKeyword(parser.Position, "xref"))
            {
                parser.Position += 4;
                return ReadClassicXref(parser);
            }

            PdfStream stream = parser.ParseIndirectObjectAt(offset) as PdfStream;
            if (stream == null) return null;
            ReadXrefStreamDictionary(stream);
            Section section = new Section();
            section.Trailer = stream.Dictionary;
            section.Previous = ReadLong(stream.Dictionary.Get("Prev"));
            return section;
        }

        /// <summary>读取经典 xref 表直至 trailer。Reads a classic xref table until its trailer.</summary>
        /// <param name="parser">位于 xref 之后的解析器。The parser positioned after xref.</param>
        /// <returns>区段信息；缺失 trailer 时返回 null。The section, or null when no trailer is present.</returns>
        private Section ReadClassicXref(PdfParser parser)
        {
            while (true)
            {
                parser.SkipWhitespaceAndComments();
                if (parser.EndOfData) return null;
                if (parser.MatchesKeyword(parser.Position, "trailer"))
                {
                    parser.Position += "trailer".Length;
                    PdfDictionary trailer = parser.ParseObject() as PdfDictionary;
                    Section section = new Section();
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
                    int number = (int)(start + i);
                    if (type == "n")
                    {
                        AddEntry(number, new XrefEntry { Type = 1, Offset = entryOffset, Generation = generation });
                    }
                    else if (type == "f")
                    {
                        AddEntry(number, new XrefEntry { Type = 0, Generation = generation });
                    }
                }
            }
        }

        /// <summary>读取 /XRefStm 指向的混合交叉表流。Reads the hybrid xref stream at /XRefStm.</summary>
        /// <param name="offset">流对象偏移。The stream object offset.</param>
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

        /// <summary>按 /W 与 /Index 解码交叉表流条目。Decodes xref stream entries using /W and /Index.</summary>
        /// <param name="stream">交叉表流。The xref stream.</param>
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

            byte[] body;
            try
            {
                body = PdfStreamDecoder.Decode(stream, Resolve);
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
                    // 先减后比，避免 position + rowLength 溢出；条目号超界直接收尾。
                    // Subtract before comparing to avoid overflow; stop when the entry number leaves int range.
                    if (position > body.Length - rowLength) return;
                    long numberLong = start + i;
                    if (numberLong > int.MaxValue) return;
                    long[] fields = new long[widthValues.Length];
                    for (int f = 0; f < widthValues.Length; f++)
                    {
                        long value = 0;
                        for (int b = 0; b < widthValues[f]; b++)
                        {
                            value = (value << 8) | body[position++];
                        }
                        fields[f] = value;
                    }
                    long type = widthValues[0] == 0 ? 1 : fields[0];
                    int number = (int)(start + i);
                    if (type == 1)
                    {
                        AddEntry(number, new XrefEntry { Type = 1, Offset = fields[1], Generation = (int)fields[2] });
                    }
                    else if (type == 2)
                    {
                        AddEntry(number, new XrefEntry
                        {
                            Type = 2,
                            Generation = 0,
                            ObjectStream = (int)fields[1],
                            ObjectIndex = (int)fields[2]
                        });
                    }
                    else if (type == 0)
                    {
                        AddEntry(number, new XrefEntry { Type = 0, Generation = (int)fields[2] });
                    }
                }
            }
        }

        /// <summary>登记交叉表条目，新条目优先。Registers an entry; newer sections win.</summary>
        /// <param name="number">对象号。The object number.</param>
        /// <param name="entry">条目。The entry.</param>
        private void AddEntry(int number, XrefEntry entry)
        {
            if (number < 0) return;
            if (entry.Type == 0 && entry.Generation == 0) return;
            if (_xref.ContainsKey(number)) return;
            // 上限兜底：畸形交叉表可以声明数百万条目，字典会无节制长大直到内存耗尽。
            // Ceiling: a malformed cross-reference table can declare millions of entries and grow
            // this dictionary until the process runs out of memory.
            if (_xref.Count >= PdfReadLimits.MaxXrefEntries) return;
            _xref[number] = entry;
            if (number > _maxObjectNumber) _maxObjectNumber = number;
        }

        /// <summary>
        /// 交叉表损坏时的救援路径：全文扫描对象头与 trailer。
        /// Recovery path for broken cross-references: scans object headers and trailers.
        /// </summary>
        private void RebuildXref()
        {
            Dictionary<int, XrefEntry> found = ScanIndirectObjects();
            foreach (KeyValuePair<int, XrefEntry> pair in found)
            {
                _xref[pair.Key] = pair.Value;
                if (pair.Key > _maxObjectNumber) _maxObjectNumber = pair.Key;
            }

            PdfDictionary trailer = FindLastTrailer();
            if (trailer == null) trailer = new PdfDictionary();
            if (trailer.Get("Root") == null)
            {
                List<int> keys = new List<int>(found.Keys);
                keys.Sort();
                for (int i = 0; i < keys.Count; i++)
                {
                    PdfDictionary dictionary = LoadObject(keys[i]) as PdfDictionary;
                    if (dictionary == null) continue;
                    PdfName type = Resolve(dictionary.Get("Type")) as PdfName;
                    if (type != null && type.Value == "Catalog")
                    {
                        trailer.Set("Root", new PdfReference(keys[i], found[keys[i]].Generation));
                        break;
                    }
                }
            }
            _trailer = trailer;
            _lastXrefOffset = 0;
        }

        /// <summary>扫描对象头并记录对象号、代次与偏移。Scans object headers and records number, generation and offset.</summary>
        /// <returns>对象号到交叉表条目的映射。A map from object number to xref entry.</returns>
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
                int generation;
                if (!int.TryParse(Encoding.ASCII.GetString(_data, numberStart, numberEnd - numberStart),
                        NumberStyles.Integer, CultureInfo.InvariantCulture, out number)) continue;
                if (!int.TryParse(Encoding.ASCII.GetString(_data, generationStart, generationEnd - generationStart),
                        NumberStyles.Integer, CultureInfo.InvariantCulture, out generation)) continue;
                if (number <= 0) continue;
                result[number] = new XrefEntry { Type = 1, Offset = numberStart, Generation = generation };
            }
            return result;
        }

        /// <summary>查找最后一个可解析的 trailer。Finds the last trailer that parses as a dictionary.</summary>
        /// <returns>trailer 字典，或 null。The trailer dictionary, or null.</returns>
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

        /// <summary>解析器回调使用的解引用入口。Dereference callback for the parser.</summary>
        /// <param name="reference">引用。The reference.</param>
        /// <returns>目标对象或 null。The target object or null.</returns>
        private PdfObject ResolveReference(PdfReference reference)
        {
            return LoadObject(reference.Number);
        }

        /// <summary>
        /// 按对象号惰性加载并解密对象，带缓存与循环保护。
        /// Lazily loads and decrypts an object by number, with caching and cycle protection.
        /// </summary>
        /// <param name="number">对象号。The object number.</param>
        /// <returns>对象；不存在或读取失败时返回 null。The object, or null when absent or unreadable.</returns>
        internal PdfObject LoadObject(int number)
        {
            if (number <= 0 || _removed.Contains(number)) return null;
            PdfObject changed;
            if (_changes.TryGetValue(number, out changed)) return changed;
            PdfObject cached;
            if (_cache.TryGetValue(number, out cached)) return cached;
            if (!_loading.Add(number)) return null;
            try
            {
                XrefEntry entry = FindEntry(number);
                PdfObject result = null;
                if (entry == null) return null;
                if (entry.Type == 1)
                {
                    // 64 位偏移直接 (int) 截断会翻成负数或落到文件中部，先做范围校验。
                    // Casting a 64-bit offset straight to int can wrap negative or land mid-file, so
                    // the offset is range-checked first.
                    if (entry.Offset < 0 || entry.Offset > _data.Length) return null;
                    PdfParser parser = new PdfParser(_data, (int)entry.Offset, ResolveReference);
                    result = parser.ParseIndirectObjectAt(entry.Offset);
                    bool encryptionDictionary = _sourceEncryptNumber > 0 && number == _sourceEncryptNumber;
                    if (result != null && _sourceSecurity != null && !encryptionDictionary)
                    {
                        DecryptObjectGraph(result, number, entry.Generation);
                    }
                }
                else if (entry.Type == 2)
                {
                    result = LoadFromObjectStream(entry.ObjectStream, entry.ObjectIndex, number);
                }
                if (result != null) _cache[number] = result;
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

        /// <summary>在交叉表或待写变更中查找对象条目。Finds an entry in the xref or pending changes.</summary>
        /// <param name="number">对象号。The object number.</param>
        /// <returns>条目或 null。The entry or null.</returns>
        private XrefEntry FindEntry(int number)
        {
            XrefEntry entry;
            if (_xref.TryGetValue(number, out entry)) return entry;
            if (_changes.ContainsKey(number)) return new XrefEntry { Type = 1, Offset = 0, Generation = 0 };
            return null;
        }

        /// <summary>从对象流中取出并解析对象（对象流自身已加密，内部字符串不再单独加密）。Extracts an object from an object stream (the stream is encrypted, its strings are not encrypted individually).</summary>
        /// <param name="streamNumber">对象流对象号。The object stream number.</param>
        /// <param name="objectIndex">流内序号。The index inside the stream.</param>
        /// <param name="targetNumber">目标对象号。The target object number.</param>
        /// <returns>对象或 null。The object or null.</returns>
        private PdfObject LoadFromObjectStream(int streamNumber, int objectIndex, int targetNumber)
        {
            PdfStream stream = LoadObject(streamNumber) as PdfStream;
            if (stream == null) return null;
            PdfDictionary dictionary = stream.Dictionary;
            int count = GetInt(dictionary.Get("N"), 0);
            int first = GetInt(dictionary.Get("First"), 0);
            byte[] body = DecodeStream(stream);
            if (body.Length == 0) return null;
            int maxObjects = Math.Min(PdfReadLimits.MaxObjectStreamObjects, body.Length / 4 + 1);
            if (count > maxObjects) count = maxObjects;

            PdfParser header = new PdfParser(body, 0, ResolveReference);
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
            int position = first + offset;
            if (position < 0 || position >= body.Length) return null;
            PdfParser parser = new PdfParser(body, position, ResolveReference);
            return parser.ParseObject();
        }

        /// <summary>
        /// 递归解密字符串与流数据；引用不在此展开。
        /// Recursively decrypts strings and stream data; references are not expanded here.
        /// </summary>
        /// <param name="value">待解密对象。The object to decrypt.</param>
        /// <param name="number">所属对象号。The owning object number.</param>
        /// <param name="generation">所属代次。The owning generation.</param>
        private void DecryptObjectGraph(PdfObject value, int number, int generation)
        {
            PdfStandardSecurityHandler security = _sourceSecurity;
            if (security == null) return;
            PdfString text = value as PdfString;
            if (text != null)
            {
                text.ReplaceBytes(security.DecryptString(text.Bytes, number, generation));
                return;
            }
            PdfArray array = value as PdfArray;
            if (array != null)
            {
                for (int i = 0; i < array.Count; i++) DecryptObjectGraph(array[i], number, generation);
                return;
            }
            PdfDictionary dictionary = value as PdfDictionary;
            if (dictionary != null)
            {
                foreach (KeyValuePair<string, PdfObject> pair in dictionary) DecryptObjectGraph(pair.Value, number, generation);
                return;
            }
            PdfStream stream = value as PdfStream;
            if (stream != null)
            {
                foreach (KeyValuePair<string, PdfObject> pair in stream.Dictionary) DecryptObjectGraph(pair.Value, number, generation);
                bool metadata = false;
                PdfName type = Resolve(stream.Dictionary.Get("Type")) as PdfName;
                if (type != null && type.Value == "Metadata") metadata = true;
                if (!(metadata && !security.EncryptMetadata))
                {
                    stream.ReplaceData(security.DecryptStream(stream.Data, number, generation));
                }
            }
        }

        /// <summary>把对象读取为整数。Reads an object as an integer.</summary>
        /// <param name="value">对象。The object.</param>
        /// <param name="fallback">默认值。The fallback.</param>
        /// <returns>整数或默认值。The integer or fallback.</returns>
        private int GetInt(PdfObject value, int fallback)
        {
            PdfNumber number = Resolve(value) as PdfNumber;
            return number == null ? fallback : (int)number.Value;
        }

        /// <summary>把对象读取为长整数。Reads an object as a long.</summary>
        /// <param name="value">对象。The object.</param>
        /// <returns>长整数或 0。The long value or 0.</returns>
        private long ReadLong(PdfObject value)
        {
            PdfNumber number = Resolve(value) as PdfNumber;
            return number == null ? 0 : (long)number.Value;
        }

        /// <summary>读取信息字典。Reads the information dictionary.</summary>
        /// <returns>元数据对象。The metadata object.</returns>
        private PdfDocumentInfo ReadInfo()
        {
            PdfDocumentInfo info = new PdfDocumentInfo();
            PdfDictionary dictionary = Resolve(_trailer == null ? null : _trailer.Get("Info")) as PdfDictionary;
            if (dictionary == null) return info;
            info.Title = ReadInfoText(dictionary.Get("Title"));
            info.Author = ReadInfoText(dictionary.Get("Author"));
            info.Subject = ReadInfoText(dictionary.Get("Subject"));
            info.Keywords = ReadInfoText(dictionary.Get("Keywords"));
            info.Creator = ReadInfoText(dictionary.Get("Creator"));
            info.Producer = ReadInfoText(dictionary.Get("Producer"));
            info.CreationDate = ParsePdfDate(ReadInfoText(dictionary.Get("CreationDate")));
            info.ModifiedDate = ParsePdfDate(ReadInfoText(dictionary.Get("ModDate")));
            return info;
        }

        /// <summary>读取并解码信息串。Reads and decodes an info string.</summary>
        /// <param name="value">对象。The object.</param>
        /// <returns>文本。The text.</returns>
        private string ReadInfoText(PdfObject value)
        {
            PdfString text = Resolve(value) as PdfString;
            return text == null ? string.Empty : PdfDocumentInfo.DecodePdfString(text.Bytes);
        }

        /// <summary>解析 PDF 日期串。Parses a PDF date string.</summary>
        /// <param name="value">日期串。The date string.</param>
        /// <returns>日期或 null。The date or null.</returns>
        private static DateTime? ParsePdfDate(string value)
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
            int year = ParseDigits(digits, 0, 4);
            int month = digits.Length >= 6 ? Math.Max(1, Math.Min(12, ParseDigits(digits, 4, 2))) : 1;
            int day = digits.Length >= 8 ? ParseDigits(digits, 6, 2) : 1;
            int hour = digits.Length >= 10 ? ParseDigits(digits, 8, 2) : 0;
            int minute = digits.Length >= 12 ? ParseDigits(digits, 10, 2) : 0;
            int second = digits.Length >= 14 ? ParseDigits(digits, 12, 2) : 0;
            try
            {
                int maxDay = DateTime.DaysInMonth(year < 1 ? 1 : year, month);
                day = Math.Max(1, Math.Min(maxDay, day));
                return new DateTime(year, month, day, Math.Min(23, hour), Math.Min(59, minute), Math.Min(59, second));
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }
        }

        /// <summary>解析定长数字。Parses a fixed-length number.</summary>
        private static int ParseDigits(string text, int start, int length)
        {
            int value = 0;
            for (int i = start; i < start + length && i < text.Length; i++)
            {
                if (text[i] >= '0' && text[i] <= '9') value = value * 10 + (text[i] - '0');
            }
            return value;
        }

        /// <summary>设置或移除文本项。Sets or removes a text entry.</summary>
        private static void SetOrRemove(PdfDictionary dictionary, string key, string value)
        {
            if (string.IsNullOrEmpty(value)) dictionary.Remove(key);
            else dictionary.SetString(key, value);
        }

        /// <summary>分配下一个对象号。Allocates the next object number.</summary>
        /// <returns>新对象号。The new object number.</returns>
        private int NextObjectNumber()
        {
            _maxObjectNumber++;
            while (_removed.Contains(_maxObjectNumber)
                || _xref.ContainsKey(_maxObjectNumber)
                || _cache.ContainsKey(_maxObjectNumber)
                || _changes.ContainsKey(_maxObjectNumber))
            {
                _maxObjectNumber++;
            }
            return _maxObjectNumber;
        }
    }
}
