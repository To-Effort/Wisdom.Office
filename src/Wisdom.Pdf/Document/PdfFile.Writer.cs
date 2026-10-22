using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Pdf
{
    /// <summary>
    /// 对象级文档的序列化：全量重写与增量更新，含加密、交叉表与 trailer。
    /// Serialization of the object-level document: full rewrites and incremental updates, including
    /// encryption, cross-reference tables and trailers.
    /// </summary>
    public sealed partial class PdfFile
    {
        /// <summary>一次写出中记录的对象偏移。An object offset recorded during one write pass.</summary>
        private sealed class WrittenObject
        {
            /// <summary>对象号。The object number.</summary>
            public int Number;
            /// <summary>代号。The generation.</summary>
            public int Generation;
            /// <summary>对象头偏移。The object header offset.</summary>
            public long Offset;
        }

        /// <summary>
        /// 执行全量或增量写出。
        /// Performs a full or incremental write.
        /// </summary>
        /// <param name="incremental">是否增量追加。Whether to append an incremental update.</param>
        /// <returns>完整 PDF 字节。The complete PDF bytes.</returns>
        private byte[] WriteFile(bool incremental)
        {
            if (_infoDirty) UpdateInfo();
            if (_isNew || !incremental) return WriteFull();

            if (_pendingEncryption != null)
            {
                throw new InvalidOperationException("Encryption settings require a full save.");
            }

            EnsureCatalogNumber();
            PdfStandardSecurityHandler security = _security;
            int encryptNumber = _encryptReference == null ? 0 : _encryptReference.Number;
            byte[] documentId = GetOrCreateDocumentId();

            SortedSet<int> changed = new SortedSet<int>();
            foreach (int number in _changes.Keys) changed.Add(number);
            foreach (int number in _removed) changed.Add(number);
            if (changed.Count == 0) return _data;

            using (MemoryStream output = new MemoryStream())
            {
                output.Write(_data, 0, _data.Length);
                if (_data.Length > 0 && _data[_data.Length - 1] != 10 && _data[_data.Length - 1] != 13)
                {
                    output.WriteByte(10);
                }

                List<WrittenObject> written = new List<WrittenObject>();
                List<int> removed = new List<int>(_removed);
                removed.Sort();

                foreach (int number in changed)
                {
                    if (_removed.Contains(number)) continue;
                    PdfObject value = LoadObject(number);
                    if (value == null) continue;
                    XrefEntry entry = FindEntry(number);
                    int generation = entry == null ? 0 : entry.Generation;
                    WrittenObject record = new WrittenObject();
                    record.Number = number;
                    record.Generation = generation;
                    record.Offset = output.Position;
                    WriteIndirectObject(output, number, generation, value, security, encryptNumber);
                    written.Add(record);
                }

                long xrefOffset = output.Position;
                WriteIncrementalXref(output, changed, written, removed);
                WriteTrailer(output, xrefOffset, true, documentId, security, encryptNumber);

                byte[] result = output.ToArray();
                CommitWriteState(result, xrefOffset, written, removed);
                return result;
            }
        }

        /// <summary>
        /// 全量重写全部对象并输出新的交叉表。
        /// Re-serializes every object and writes a fresh cross-reference table.
        /// </summary>
        /// <returns>完整 PDF 字节。The complete PDF bytes.</returns>
        private byte[] WriteFull()
        {
            EnsureCatalogNumber();
            byte[] documentId = GetOrCreateDocumentId();
            PdfStandardSecurityHandler security = _security;
            int encryptNumber = _encryptReference == null ? 0 : _encryptReference.Number;

            if (_pendingEncryption != null)
            {
                security = PdfStandardSecurityHandler.Create(_pendingEncryption, documentId);
                encryptNumber = AddObject(security.EncryptionDictionary);
            }

            SortedDictionary<int, PdfObject> objects = new SortedDictionary<int, PdfObject>();
            foreach (int number in ObjectNumbers)
            {
                if (number == encryptNumber || number == _sourceEncryptNumber) continue;
                PdfObject value = LoadObject(number);
                if (value != null) objects[number] = value;
            }
            if (security != null && encryptNumber > 0)
            {
                objects[encryptNumber] = security.EncryptionDictionary;
            }
            objects[_catalogNumber] = Catalog;

            using (MemoryStream output = new MemoryStream())
            {
                WriteAscii(output, "%PDF-" + _version + "\n");
                output.Write(new byte[] { 0x25, 0xE2, 0xE3, 0xCF, 0xD3, 0x0A }, 0, 6);

                List<WrittenObject> written = new List<WrittenObject>();
                foreach (KeyValuePair<int, PdfObject> pair in objects)
                {
                    XrefEntry entry;
                    int generation = _xref.TryGetValue(pair.Key, out entry) ? entry.Generation : 0;
                    WrittenObject record = new WrittenObject();
                    record.Number = pair.Key;
                    record.Generation = generation;
                    record.Offset = output.Position;
                    WriteIndirectObject(output, pair.Key, generation, pair.Value, security, encryptNumber);
                    written.Add(record);
                }

                long xrefOffset = output.Position;
                int maxNumber = 0;
                for (int i = 0; i < written.Count; i++) if (written[i].Number > maxNumber) maxNumber = written[i].Number;
                if (_maxObjectNumber > maxNumber) maxNumber = _maxObjectNumber;
                WriteAscii(output, "xref\n0 " + (maxNumber + 1).ToString(CultureInfo.InvariantCulture) + "\n");
                WriteAscii(output, "0000000000 65535 f\r\n");
                int cursor = 1;
                for (int i = 0; i < written.Count; i++)
                {
                    WrittenObject record = written[i];
                    while (cursor < record.Number)
                    {
                        WriteAscii(output, "0000000000 00000 f\r\n");
                        cursor++;
                    }
                    WriteAscii(output, record.Offset.ToString("0000000000", CultureInfo.InvariantCulture)
                        + " " + record.Generation.ToString("00000", CultureInfo.InvariantCulture) + " n\r\n");
                    cursor++;
                }
                while (cursor <= maxNumber)
                {
                    WriteAscii(output, "0000000000 00000 f\r\n");
                    cursor++;
                }

                WriteTrailer(output, xrefOffset, false, documentId, security, encryptNumber);

                byte[] result = output.ToArray();
                CommitWriteState(result, xrefOffset, written, null);
                ApplySecurityState(security, encryptNumber);
                return result;
            }
        }

        /// <summary>
        /// 写出一个间接对象（对象头、对象体、endobj）。
        /// Writes one indirect object: header, body and endobj.
        /// </summary>
        /// <param name="output">输出流。The output stream.</param>
        /// <param name="number">对象号。The object number.</param>
        /// <param name="generation">代号。The generation.</param>
        /// <param name="value">对象。The object.</param>
        /// <param name="security">加密处理器，可为 null。The security handler, or null.</param>
        /// <param name="encryptNumber">加密字典的对象号。The object number of the encryption dictionary.</param>
        private static void WriteIndirectObject(Stream output, int number, int generation, PdfObject value,
            PdfStandardSecurityHandler security, int encryptNumber)
        {
            WriteAscii(output, number.ToString(CultureInfo.InvariantCulture) + " "
                + generation.ToString(CultureInfo.InvariantCulture) + " obj\n");
            WriteObject(output, value, security, number, generation, encryptNumber);
            WriteAscii(output, "\nendobj\n");
        }

        /// <summary>
        /// 递归写出对象语法；字符串与流按需加密。
        /// Recursively writes PDF syntax and encrypts strings and streams on demand.
        /// </summary>
        /// <param name="output">输出流。The output stream.</param>
        /// <param name="value">对象。The object.</param>
        /// <param name="security">加密处理器。The security handler.</param>
        /// <param name="number">所属对象号。The owning object number.</param>
        /// <param name="generation">所属代次。The owning generation.</param>
        /// <param name="encryptNumber">加密字典对象号。The encryption dictionary number.</param>
        private static void WriteObject(Stream output, PdfObject value, PdfStandardSecurityHandler security,
            int number, int generation, int encryptNumber)
        {
            if (value == null || value is PdfNull)
            {
                WriteAscii(output, "null");
                return;
            }
            PdfBoolean boolean = value as PdfBoolean;
            if (boolean != null)
            {
                WriteAscii(output, boolean.Value ? "true" : "false");
                return;
            }
            PdfNumber numberValue = value as PdfNumber;
            if (numberValue != null)
            {
                WriteAscii(output, PdfNumberFormat.Format(numberValue.Value));
                return;
            }
            PdfName name = value as PdfName;
            if (name != null)
            {
                WriteAscii(output, "/" + PdfNameEscape.Escape(name.Value));
                return;
            }
            PdfString text = value as PdfString;
            if (text != null)
            {
                byte[] bytes = text.Bytes;
                if (security != null && number != encryptNumber)
                {
                    bytes = security.EncryptString(bytes, number, generation);
                }
                WriteHexString(output, bytes);
                return;
            }
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
                    WriteObject(output, array[i], security, number, generation, encryptNumber);
                }
                WriteAscii(output, "]");
                return;
            }
            PdfDictionary dictionary = value as PdfDictionary;
            if (dictionary != null)
            {
                WriteDictionary(output, dictionary, security, number, generation, encryptNumber);
                return;
            }
            PdfStream stream = value as PdfStream;
            if (stream != null)
            {
                PdfDictionary streamDictionary = stream.Dictionary.Clone();
                byte[] data = stream.Data;
                if (stream.Compress)
                {
                    data = Zlib.Compress(data);
                    streamDictionary.Remove("Filter");
                    streamDictionary.SetName("Filter", "FlateDecode");
                    streamDictionary.Remove("DecodeParms");
                }
                if (security != null && number != encryptNumber)
                {
                    bool metadata = !security.EncryptMetadata
                        && PdfStandardSecurityHandler.IsMetadataStream(streamDictionary);
                    if (!metadata) data = security.EncryptStream(data, number, generation);
                }
                streamDictionary.SetNumber("Length", data.Length);
                WriteDictionary(output, streamDictionary, security, number, generation, encryptNumber);
                WriteAscii(output, "\nstream\n");
                output.Write(data, 0, data.Length);
                WriteAscii(output, "\nendstream");
                return;
            }
            throw new NotSupportedException("Unsupported PDF object type: " + value.GetType().FullName);
        }

        /// <summary>写出字典。Writes a dictionary.</summary>
        private static void WriteDictionary(Stream output, PdfDictionary dictionary, PdfStandardSecurityHandler security,
            int number, int generation, int encryptNumber)
        {
            WriteAscii(output, "<<");
            bool first = true;
            foreach (KeyValuePair<string, PdfObject> pair in dictionary)
            {
                if (!first) WriteAscii(output, " ");
                first = false;
                WriteAscii(output, "/" + PdfNameEscape.Escape(pair.Key) + " ");
                WriteObject(output, pair.Value, security, number, generation, encryptNumber);
            }
            WriteAscii(output, ">>");
        }

        /// <summary>写出十六进制字符串。Writes a hexadecimal string.</summary>
        /// <param name="output">输出流。The output stream.</param>
        /// <param name="bytes">字节内容。The bytes.</param>
        private static void WriteHexString(Stream output, byte[] bytes)
        {
            output.WriteByte((byte)'<');
            for (int i = 0; i < bytes.Length; i++)
            {
                WriteAscii(output, bytes[i].ToString("x2", CultureInfo.InvariantCulture));
            }
            output.WriteByte((byte)'>');
        }

        /// <summary>
        /// 写出增量交叉表：0 号空闲条目加若干连续区段。
        /// Writes an incremental cross-reference table: the free object 0 entry plus contiguous subsections.
        /// </summary>
        /// <param name="output">输出流。The output stream.</param>
        /// <param name="changed">变更与删除的对象号。Changed and removed object numbers.</param>
        /// <param name="written">已写对象记录。Written object records.</param>
        /// <param name="removed">已删除对象号。Removed object numbers.</param>
        private void WriteIncrementalXref(Stream output, SortedSet<int> changed, List<WrittenObject> written, List<int> removed)
        {
            WriteAscii(output, "xref\n0 1\n0000000000 65535 f\r\n");

            Dictionary<int, WrittenObject> index = new Dictionary<int, WrittenObject>();
            for (int i = 0; i < written.Count; i++) index[written[i].Number] = written[i];

            List<int> numbers = new List<int>();
            foreach (int number in changed)
            {
                if (index.ContainsKey(number) || removed.Contains(number)) numbers.Add(number);
            }
            numbers.Sort();

            int position = 0;
            while (position < numbers.Count)
            {
                int start = numbers[position];
                int end = start;
                while (position + 1 < numbers.Count && numbers[position + 1] == end + 1)
                {
                    position++;
                    end = numbers[position];
                }
                WriteAscii(output, start.ToString(CultureInfo.InvariantCulture) + " "
                    + (end - start + 1).ToString(CultureInfo.InvariantCulture) + "\n");
                for (int number = start; number <= end; number++)
                {
                    WrittenObject record;
                    if (index.TryGetValue(number, out record))
                    {
                        WriteAscii(output, record.Offset.ToString("0000000000", CultureInfo.InvariantCulture)
                            + " " + record.Generation.ToString("00000", CultureInfo.InvariantCulture) + " n\r\n");
                    }
                    else
                    {
                        XrefEntry entry;
                        int generation = _xref.TryGetValue(number, out entry) ? entry.Generation + 1 : 1;
                        WriteAscii(output, "0000000000 " + generation.ToString("00000", CultureInfo.InvariantCulture) + " f\r\n");
                    }
                }
                position++;
            }
        }

        /// <summary>
        /// 写出 trailer 与 startxref。
        /// Writes the trailer and startxref.
        /// </summary>
        /// <param name="output">输出流。The output stream.</param>
        /// <param name="xrefOffset">交叉表偏移。The xref offset.</param>
        /// <param name="incremental">是否增量更新。Whether this is an incremental update.</param>
        /// <param name="documentId">文档标识。The document identifier.</param>
        /// <param name="security">加密处理器。The security handler.</param>
        /// <param name="encryptNumber">加密字典对象号。The encryption dictionary number.</param>
        private void WriteTrailer(Stream output, long xrefOffset, bool incremental, byte[] documentId,
            PdfStandardSecurityHandler security, int encryptNumber)
        {
            int maxNumber = _maxObjectNumber;
            foreach (int number in _changes.Keys) if (number > maxNumber) maxNumber = number;
            foreach (int number in _xref.Keys) if (number > maxNumber) maxNumber = number;

            PdfDictionary trailer = new PdfDictionary();
            trailer.SetNumber("Size", maxNumber + 1);
            trailer.Set("Root", new PdfReference(_catalogNumber, 0));
            PdfObject info = _trailer == null ? null : _trailer.Get("Info");
            if (info != null) trailer.Set("Info", info);
            if (incremental && _lastXrefOffset > 0) trailer.SetNumber("Prev", _lastXrefOffset);
            PdfArray id = new PdfArray();
            id.Add(PdfString.Hex(documentId));
            id.Add(PdfString.Hex(documentId));
            trailer.Set("ID", id);
            if (security != null && encryptNumber > 0) trailer.Set("Encrypt", new PdfReference(encryptNumber, 0));

            WriteAscii(output, "trailer\n");
            WriteObject(output, trailer, null, 0, 0, 0);
            WriteAscii(output, "\nstartxref\n" + xrefOffset.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
        }

        /// <summary>
        /// 取得或生成 /ID 的第一个标识串；缺少时生成 16 字节随机标识并写回 trailer。
        /// Gets or creates the first /ID string; a 16-byte random identifier is generated and stored when absent.
        /// </summary>
        /// <returns>标识字节。The identifier bytes.</returns>
        private byte[] GetOrCreateDocumentId()
        {
            if (_trailer == null) _trailer = new PdfDictionary();
            PdfArray id = Resolve(_trailer.Get("ID")) as PdfArray;
            if (id != null && id.Count > 0)
            {
                PdfString existing = Resolve(id[0]) as PdfString;
                if (existing != null && existing.Bytes.Length > 0) return existing.Bytes;
            }
            byte[] bytes = new byte[16];
            using (System.Security.Cryptography.RandomNumberGenerator random = System.Security.Cryptography.RandomNumberGenerator.Create())
            {
                random.GetBytes(bytes);
            }
            PdfArray replacement = new PdfArray();
            replacement.Add(PdfString.Hex(bytes));
            replacement.Add(PdfString.Hex(bytes));
            _trailer.Set("ID", replacement);
            return bytes;
        }

        /// <summary>确保目录被登记为间接对象并更新 trailer。Ensures the catalog is an indirect object and updates the trailer.</summary>
        private void EnsureCatalogNumber()
        {
            PdfDictionary catalog = Catalog;
            if (_catalogNumber <= 0)
            {
                _catalogNumber = AddObject(catalog);
            }
            if (_trailer == null) _trailer = new PdfDictionary();
            _trailer.Set("Root", new PdfReference(_catalogNumber, 0));
            if (_trailer.Get("Info") == null)
            {
                PdfReference info = GetOrCreateInfoReference();
                _trailer.Set("Info", info);
            }
        }

        /// <summary>
        /// 全量保存成功后同步加密状态：输出处理器同时成为后续读取源对象的处理器。
        /// Synchronizes security state after a successful full save: the output handler also becomes the
        /// handler used to read source objects afterwards.
        /// </summary>
        /// <param name="security">输出使用的处理器，未加密时为 null。The handler used for output, or null when unencrypted.</param>
        /// <param name="encryptNumber">输出的加密字典对象号，未加密时为 0。The output encryption dictionary number, or 0.</param>
        private void ApplySecurityState(PdfStandardSecurityHandler security, int encryptNumber)
        {
            _security = security;
            _sourceSecurity = security;
            _sourceEncryptNumber = security == null ? 0 : encryptNumber;
            _encryptReference = security == null ? null : new PdfReference(encryptNumber, 0);
            _pendingEncryption = null;
            _rewriteRequired = false;
            if (_trailer != null)
            {
                if (security != null) _trailer.Set("Encrypt", _encryptReference);
                else _trailer.Remove("Encrypt");
            }
        }

        /// <summary>
        /// 保存后更新内存状态，使后续增量保存可以继续。
        /// Updates the in-memory state after saving so later incremental saves can continue.
        /// </summary>
        /// <param name="bytes">完整文件字节。The complete file bytes.</param>
        /// <param name="xrefOffset">本次交叉表偏移。The xref offset of this write.</param>
        /// <param name="written">已写对象。Written objects.</param>
        /// <param name="removed">删除对象。Removed objects.</param>
        private void CommitWriteState(byte[] bytes, long xrefOffset, List<WrittenObject> written, List<int> removed)
        {
            _data = bytes;
            _lastXrefOffset = xrefOffset;
            _isNew = false;
            for (int i = 0; i < written.Count; i++)
            {
                WrittenObject record = written[i];
                _xref[record.Number] = new XrefEntry { Type = 1, Offset = record.Offset, Generation = record.Generation };
                if (record.Number > _maxObjectNumber) _maxObjectNumber = record.Number;
            }
            if (removed != null)
            {
                for (int i = 0; i < removed.Count; i++)
                {
                    int number = removed[i];
                    XrefEntry entry;
                    int generation = _xref.TryGetValue(number, out entry) ? entry.Generation + 1 : 1;
                    _xref[number] = new XrefEntry { Type = 0, Generation = generation };
                }
            }
            _changes.Clear();
            _removed.Clear();
            _pageNodes = null;
            _trailer = BuildCurrentTrailer(_trailer, xrefOffset);
        }

        /// <summary>构造保存后应保留的 trailer（避免把临时 /Prev 留在内存）。Builds the trailer to retain in memory.</summary>
        private PdfDictionary BuildCurrentTrailer(PdfDictionary previous, long xrefOffset)
        {
            PdfDictionary trailer = previous == null ? new PdfDictionary() : previous.Clone();
            trailer.Remove("Prev");
            trailer.Remove("XRefStm");
            trailer.SetNumber("Size", _maxObjectNumber + 1);
            trailer.Set("Root", new PdfReference(_catalogNumber, 0));
            return trailer;
        }

        /// <summary>以 ASCII 写出一段结构文本。Writes structural text as ASCII.</summary>
        /// <param name="output">输出流。The output stream.</param>
        /// <param name="text">文本。The text.</param>
        private static void WriteAscii(Stream output, string text)
        {
            byte[] bytes = Encoding.ASCII.GetBytes(text);
            output.Write(bytes, 0, bytes.Length);
        }
    }
}
