using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Wisdom.Pdf;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// 标准安全处理器（RC4-40 / RC4-128 / AES-128 / AES-256）的加密写出与口令读回回归测试。
    /// Regression tests for the standard security handler: writing RC4-40, RC4-128, AES-128 and
    /// AES-256 documents and re-opening them with either password.
    /// </summary>
    /// <remarks>
    /// 除了自洽回读，本文件内嵌一份按 ISO 32000-2 算法 2.B/3.8/3.9/3.10 独立实现的参考散列，用它校验库写出的
    /// <c>/U</c>、<c>/O</c>、<c>/UE</c>、<c>/Perms</c>，因此能挡住"自己写得对、别人读不懂"的自洽错误；
    /// 另一条用例把 <c>/P</c> 改写成无符号形式（pypdf 等写出器的实际做法），确认读取端不会因 32 位溢出而算错文件密钥。
    /// Besides self-consistent round trips, the file embeds a reference implementation of the ISO 32000-2
    /// algorithms 2.B/3.8/3.9/3.10 and uses it to check the <c>/U</c>, <c>/O</c>, <c>/UE</c> and <c>/Perms</c>
    /// the library writes, which catches self-consistent-but-wrong output that other readers would reject.
    /// A second case rewrites <c>/P</c> in unsigned form (what pypdf and similar writers emit) to prove the
    /// reader does not derive a wrong file key through 32-bit overflow.
    /// </remarks>
    internal static partial class Program
    {
        /// <summary>用户口令。The user password.</summary>
        private const string EncUserPassword = "enc-user";
        /// <summary>所有者口令。The owner password.</summary>
        private const string EncOwnerPassword = "enc-owner";

        private static void RunPdfEncryptionTest(string sourcePdfPath, string outputDirectory)
        {
            byte[] source = File.ReadAllBytes(sourcePdfPath);

            RunEncryptionRoundTrip(source, outputDirectory);
            RunR2ToR4MatchesStandardVector(source, outputDirectory);
            RunAes256MatchesStandardVector(source, outputDirectory);
            RunUnsignedPermissionsStillOpens(source);
            RunAes256Revision5Fixture(outputDirectory);

            Console.WriteLine("PDF encryption tests passed.");
        }

        /// <summary>用独立实现的算法 2/4/5 校验 RC4-40、RC4-128 与 AES-128 写出的 /O、/U，覆盖元数据明文开关。
        /// Validates the written /O and /U of RC4-40, RC4-128 and AES-128 against an independent implementation of algorithms 2, 4 and 5, including the metadata-in-plain-text switch.</summary>
        /// <param name="source">无加密的源 PDF 字节。The unencrypted source PDF bytes.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunR2ToR4MatchesStandardVector(byte[] source, string outputDirectory)
        {
            PdfEncryptionAlgorithm[] algorithms =
            {
                PdfEncryptionAlgorithm.Rc4_40,
                PdfEncryptionAlgorithm.Rc4_128,
                PdfEncryptionAlgorithm.Aes128
            };

            foreach (PdfEncryptionAlgorithm algorithm in algorithms)
            {
                foreach (bool encryptMetadata in new bool[] { true, false })
                {
                    PdfFile file = PdfFile.Load(source);
                    PdfEncryptionOptions options = new PdfEncryptionOptions();
                    options.Algorithm = algorithm;
                    options.UserPassword = EncUserPassword;
                    options.OwnerPassword = EncOwnerPassword;
                    options.Permissions = PdfPermissions.Print | PdfPermissions.Copy;
                    options.EncryptMetadata = encryptMetadata;
                    file.SetEncryption(options);

                    string tag = algorithm + (encryptMetadata ? "" : "-plain-meta");
                    string path = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-enc-" + tag + ".pdf"));
                    File.WriteAllBytes(path, file.SaveToBytes());

                    PdfFile opened = PdfFile.Load(path, EncUserPassword);
                    PdfDictionary encrypt = EncryptionDictionary(opened);
                    int revision = (int)((PdfNumber)encrypt.Get("R")).Value;
                    int keyBits = algorithm == PdfEncryptionAlgorithm.Rc4_40 ? 40 : 128;
                    Require(revision == (algorithm == PdfEncryptionAlgorithm.Rc4_40 ? 2
                        : algorithm == PdfEncryptionAlgorithm.Rc4_128 ? 3 : 4), tag + ": unexpected /R " + revision);

                    byte[] o = StringBytes(encrypt, "O");
                    byte[] u = StringBytes(encrypt, "U");
                    byte[] id = FirstDocumentId(opened);
                    int p = unchecked((int)(long)Math.Round(((PdfNumber)encrypt.Get("P")).Value));

                    byte[] key = ReferenceFileKey(PadPassword(EncUserPassword), o, p, id, revision, keyBits, encryptMetadata);
                    byte[] expected = revision == 2 ? Rc4(key, PasswordPaddingValue)
                        : Rc4Rounds(key, Md5(Concat(PasswordPaddingValue, id)));
                    Require(BytesEqual(Slice(expected, 0, 16), Slice(u, 0, 16)),
                        tag + ": /U does not match the algorithm " + (revision == 2 ? "4" : "5") + " value of the user password");

                    byte[] ownerPadded = PadPassword(EncOwnerPassword);
                    byte[] recovered = ReferenceOwnerPassword(o, ownerPadded, revision, keyBits);
                    byte[] ownerKey = ReferenceFileKey(recovered, o, p, id, revision, keyBits, encryptMetadata);
                    byte[] ownerExpected = revision == 2 ? Rc4(ownerKey, PasswordPaddingValue)
                        : Rc4Rounds(ownerKey, Md5(Concat(PasswordPaddingValue, id)));
                    Require(BytesEqual(Slice(ownerExpected, 0, 16), Slice(u, 0, 16)),
                        tag + ": the owner password must recover the same file key");

                    if (algorithm == PdfEncryptionAlgorithm.Aes128)
                    {
                        PdfDictionary filters = encrypt.Get("CF") as PdfDictionary;
                        Require(filters != null, tag + ": AES-128 must declare /CF");
                        PdfDictionary standard = filters.Get("StdCF") as PdfDictionary;
                        Require(standard != null && ((PdfName)standard.Get("CFM")).Value == "AESV2",
                            tag + ": /StdCF must be /AESV2");
                        Require(((PdfNumber)standard.Get("Length")).Value == 16, tag + ": /AESV2 length must be 16 bytes");
                        Require(((PdfName)encrypt.Get("StmF")).Value == "StdCF" && ((PdfName)encrypt.Get("StrF")).Value == "StdCF",
                            tag + ": /StmF and /StrF must point at /StdCF");
                        Require(((PdfName)standard.Get("AuthEvent")).Value == "DocOpen", tag + ": /AuthEvent must be DocOpen");
                    }

                    RequireContains(PdfDocument.Load(path, EncOwnerPassword).ExtractText(), "Sales Report");
                }
            }
        }

        /// <summary>读取外部工具写出的 AES-256 R5 文档：<c>/U</c> 只用 SHA-256，没有算法 2.B 的迭代加固。
        /// Reads an AES-256 R5 document written by an external tool: <c>/U</c> is plain SHA-256 without the iterated Algorithm 2.B.</summary>
        /// <param name="outputDirectory">输出目录，其上一级的 fixtures 子目录存放样例。The output directory; the sibling fixtures directory holds the sample.</param>
        private static void RunAes256Revision5Fixture(string outputDirectory)
        {
            string fixture = Path.GetFullPath(Path.Combine(outputDirectory, "..", "fixtures", "pdf-aes256-r5.pdf"));
            Require(File.Exists(fixture), "the AES-256 R5 fixture is missing: " + fixture);

            PdfFile file = PdfFile.Load(fixture, "r5-user");
            Require(file.EncryptionAlgorithm == "AES-256 (R5)",
                "the fixture must report revision 5, got " + file.EncryptionAlgorithm);
            RequireContains(PdfDocument.Load(fixture, "r5-user").ExtractText(), "R5FIXTURE");

            PdfFile owner = PdfFile.Load(fixture, "r5-owner");
            Require(owner.PageCount == 1, "the owner password must open the R5 fixture");

            RequireThrows(delegate { PdfFile.Load(fixture, "r5-wrong"); }, "a wrong password must not open the R5 fixture");
        }

        /// <summary>四种算法都要能用用户口令与所有者口令读回，写出的 <c>/P</c> 必须是有符号形式，错口令必须被拒。
        /// All four algorithms must re-open under either password, the written <c>/P</c> must stay signed, and a wrong password must be rejected.</summary>
        /// <param name="source">无加密的源 PDF 字节。The unencrypted source PDF bytes.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunEncryptionRoundTrip(byte[] source, string outputDirectory)
        {
            PdfEncryptionAlgorithm[] algorithms =
            {
                PdfEncryptionAlgorithm.Rc4_40,
                PdfEncryptionAlgorithm.Rc4_128,
                PdfEncryptionAlgorithm.Aes128,
                PdfEncryptionAlgorithm.Aes256
            };

            foreach (PdfEncryptionAlgorithm algorithm in algorithms)
            {
                PdfFile file = PdfFile.Load(source);
                int pages = file.PageCount;

                PdfEncryptionOptions options = new PdfEncryptionOptions();
                options.Algorithm = algorithm;
                options.UserPassword = EncUserPassword;
                options.OwnerPassword = EncOwnerPassword;
                options.Permissions = PdfPermissions.Print | PdfPermissions.Copy;
                file.SetEncryption(options);

                string path = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-enc-" + algorithm.ToString() + ".pdf"));
                File.WriteAllBytes(path, file.SaveToBytes());

                PdfFile asUser = PdfFile.Load(path, EncUserPassword);
                Require(asUser.PageCount == pages, algorithm + ": page count changed through encryption");
                Require(asUser.IsEncrypted, algorithm + ": the re-opened document must still report encryption");
                RequireContains(PdfDocument.Load(path, EncUserPassword).ExtractText(), "Sales Report");

                PdfFile asOwner = PdfFile.Load(path, EncOwnerPassword);
                Require(asOwner.PageCount == pages, algorithm + ": the owner password did not open the document");
                Require(asOwner.EncryptionAlgorithm == asUser.EncryptionAlgorithm,
                    algorithm + ": the two passwords must report the same algorithm");

                RequireThrows(delegate { PdfFile.Load(path, "wrong-password"); },
                    algorithm + ": a wrong password must be rejected");

                PdfDictionary encrypt = EncryptionDictionary(asUser);
                PdfNumber p = encrypt.Get("P") as PdfNumber;
                Require(p != null && p.Value < 0, algorithm + ": /P must be written as a signed 32-bit value, got " +
                    (p == null ? "null" : p.Value.ToString()));
            }
        }

        /// <summary>用独立实现的算法 2.B 校验 AES-256 写出的 /U、/O、/UE 与 /Perms，确保第三方阅读器能口令开箱。
        /// Validates the AES-256 /U, /O, /UE and /Perms against an independent Algorithm 2.B so that third-party readers can open the file by password.</summary>
        /// <param name="source">无加密的源 PDF 字节。The unencrypted source PDF bytes.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunAes256MatchesStandardVector(byte[] source, string outputDirectory)
        {
            PdfFile file = PdfFile.Load(source);
            PdfEncryptionOptions options = new PdfEncryptionOptions();
            options.Algorithm = PdfEncryptionAlgorithm.Aes256;
            options.UserPassword = EncUserPassword;
            options.OwnerPassword = EncOwnerPassword;
            options.Permissions = PdfPermissions.All;
            file.SetEncryption(options);

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-enc-aes256-spec.pdf"));
            File.WriteAllBytes(path, file.SaveToBytes());

            PdfFile opened = PdfFile.Load(path, EncUserPassword);
            Require(opened.EncryptionAlgorithm == "AES-256 (R6)",
                "AES-256 must be written as revision 6, got " + opened.EncryptionAlgorithm);

            PdfDictionary encrypt = EncryptionDictionary(opened);
            byte[] u = StringBytes(encrypt, "U");
            byte[] o = StringBytes(encrypt, "O");
            byte[] ue = StringBytes(encrypt, "UE");
            byte[] perms = StringBytes(encrypt, "Perms");
            Require(u.Length == 48, "/U must be 48 bytes for AES-256, got " + u.Length);
            Require(o.Length == 48, "/O must be 48 bytes for AES-256, got " + o.Length);
            Require(ue.Length == 32, "/UE must be 32 bytes for AES-256, got " + ue.Length);
            Require(perms.Length == 16, "/Perms must be 16 bytes for AES-256, got " + perms.Length);

            byte[] userPassword = Encoding.UTF8.GetBytes(EncUserPassword);
            byte[] ownerPassword = Encoding.UTF8.GetBytes(EncOwnerPassword);
            byte[] userValue = Slice(u, 0, 48);

            Require(BytesEqual(ReferenceHashR6(userPassword, Slice(u, 32, 8), new byte[0]), Slice(u, 0, 32)),
                "/U does not match the algorithm 2.B hash of the user password");
            Require(BytesEqual(ReferenceHashR6(ownerPassword, Slice(o, 32, 8), userValue), Slice(o, 0, 32)),
                "/O does not match the algorithm 2.B hash of the owner password");

            byte[] fileKey = AesCbc(ReferenceHashR6(userPassword, Slice(u, 40, 8), new byte[0]),
                new byte[16], ue, false);
            Require(fileKey.Length == 32, "the recovered file key must be 32 bytes, got " + fileKey.Length);
            Require(BytesEqual(AesCbc(ReferenceHashR6(ownerPassword, Slice(o, 40, 8), userValue),
                new byte[16], StringBytes(encrypt, "OE"), false), fileKey),
                "/OE must recover the same file key as /UE");

            byte[] plain = AesEcb(fileKey, perms, false);
            int p = unchecked((int)(long)Math.Round(((PdfNumber)encrypt.Get("P")).Value));
            Require(plain[0] == (byte)(p & 0xFF) && plain[1] == (byte)((p >> 8) & 0xFF) &&
                plain[2] == (byte)((p >> 16) & 0xFF) && plain[3] == (byte)((p >> 24) & 0xFF),
                "/Perms must carry the /P value little-endian first");
            Require(plain[4] == 0xFF && plain[5] == 0xFF && plain[6] == 0xFF && plain[7] == 0xFF,
                "/Perms bytes 4-7 must be 0xFFFFFFFF");
            Require(plain[8] == (byte)'T', "/Perms byte 8 must be 'T' when metadata is encrypted");
            Require(plain[9] == (byte)'a' && plain[10] == (byte)'d' && plain[11] == (byte)'b',
                "/Perms bytes 9-11 must be 'adb'");

            RequireContains(PdfDocument.Load(path, EncUserPassword).ExtractText(), "Sales Report");
        }

        /// <summary>把写出的 <c>/P</c> 改写成无符号整数（pypdf 风格）后，读取端仍要能算出同一文件密钥。
        /// After rewriting the stored <c>/P</c> as an unsigned integer (pypdf style) the reader must still derive the same file key.</summary>
        /// <param name="source">无加密的源 PDF 字节。The unencrypted source PDF bytes.</param>
        private static void RunUnsignedPermissionsStillOpens(byte[] source)
        {
            PdfFile file = PdfFile.Load(source);
            int pages = file.PageCount;

            PdfEncryptionOptions options = new PdfEncryptionOptions();
            options.Algorithm = PdfEncryptionAlgorithm.Rc4_128;
            options.UserPassword = EncUserPassword;
            options.OwnerPassword = EncOwnerPassword;
            options.Permissions = PdfPermissions.All;
            file.SetEncryption(options);
            byte[] signed = file.SaveToBytes();

            PdfFile carrier = PdfFile.Load(signed, EncOwnerPassword);
            int number = EncryptObjectNumber(carrier);
            PdfDictionary encrypt = carrier.GetObject<PdfDictionary>(number);
            PdfNumber original = (PdfNumber)encrypt.Get("P");
            Require(original != null && original.Value < 0, "the writer must emit a signed /P before this test rewrites it");

            encrypt.Set("P", new PdfNumber(unchecked((uint)(int)Math.Round(original.Value))));
            carrier.MarkModified(number);
            byte[] unsigned = carrier.SaveToBytes();

            PdfDocument document = PdfDocument.Load(unsigned, EncUserPassword);
            RequireContains(document.ExtractText(), "Sales Report");
            Require(document.PageCount == pages, "the unsigned /P document must keep its page count");
        }

        /// <summary>取已打开文档的加密字典。Returns the encryption dictionary of an opened document.</summary>
        /// <param name="file">已用口令打开的文档。The document opened with a password.</param>
        /// <returns>加密字典。The encryption dictionary.</returns>
        private static PdfDictionary EncryptionDictionary(PdfFile file)
        {
            return file.GetObject<PdfDictionary>(EncryptObjectNumber(file));
        }

        /// <summary>取 trailer 中 /Encrypt 的对象号。Returns the object number of the trailer /Encrypt entry.</summary>
        /// <param name="file">已用口令打开的文档。The document opened with a password.</param>
        /// <returns>对象号。The object number.</returns>
        private static int EncryptObjectNumber(PdfFile file)
        {
            PdfReference reference = file.Trailer.Get("Encrypt") as PdfReference;
            Require(reference != null, "the saved document must carry an /Encrypt reference in the trailer");
            return reference.Number;
        }

        /// <summary>读加密字典里某个串对象的原始字节。Reads the raw bytes of a string entry in the encryption dictionary.</summary>
        /// <param name="dictionary">加密字典。The encryption dictionary.</param>
        /// <param name="key">条目名。The entry name.</param>
        /// <returns>原始字节。The raw bytes.</returns>
        private static byte[] StringBytes(PdfDictionary dictionary, string key)
        {
            PdfString text = dictionary.Get(key) as PdfString;
            Require(text != null, "the encryption dictionary is missing /" + key);
            return text.Bytes;
        }

        /// <summary>算法 2.B：SHA-256 起算，按密文前 16 字节之和选散列，至少 64 轮，取前 32 字节。
        /// Algorithm 2.B: start from SHA-256, pick the hash from the modular sum of the first 16 ciphertext
        /// bytes, run at least 64 rounds and return the first 32 bytes.</summary>
        /// <param name="password">UTF-8 口令。The UTF-8 password.</param>
        /// <param name="salt">8 字节盐。The 8-byte salt.</param>
        /// <param name="udata">所有者口令时为 48 字节 /U，否则为空。The 48-byte /U for the owner password, empty otherwise.</param>
        /// <returns>32 字节散列。The 32-byte hash.</returns>
        private static byte[] ReferenceHashR6(byte[] password, byte[] salt, byte[] udata)
        {
            byte[] k = Sha(1, Concat(password, salt, udata));
            int count = 0;
            while (true)
            {
                byte[] seed = Concat(password, k, udata);
                byte[] block = new byte[seed.Length * 64];
                for (int i = 0; i < 64; i++) Buffer.BlockCopy(seed, 0, block, i * seed.Length, seed.Length);

                byte[] e = AesCbc(Slice(k, 0, 16), Slice(k, 16, 16), block, true);
                int sum = 0;
                for (int i = 0; i < 16; i++) sum += e[i];
                k = Sha((sum % 3) + 1, e);

                count++;
                if (count >= 64 && e[e.Length - 1] <= count - 32) break;
            }
            return Slice(k, 0, 32);
        }

        /// <summary>按 1/2/3 选择 SHA-256/384/512。Selects SHA-256, SHA-384 or SHA-512 from 1, 2, 3.</summary>
        /// <param name="kind">1 为 SHA-256，2 为 SHA-384，3 为 SHA-512。1 is SHA-256, 2 SHA-384, 3 SHA-512.</param>
        /// <param name="data">输入。The input.</param>
        /// <returns>摘要。The digest.</returns>
        private static byte[] Sha(int kind, byte[] data)
        {
            using (HashAlgorithm hash = kind == 2 ? (HashAlgorithm)SHA384.Create() : kind == 3 ? SHA512.Create() : SHA256.Create())
            {
                return hash.ComputeHash(data);
            }
        }

        /// <summary>AES-CBC 无填充加解密。AES-CBC without padding, for encryption and decryption.</summary>
        /// <param name="key">密钥（16 或 32 字节）。The 16- or 32-byte key.</param>
        /// <param name="iv">初始向量。The initialization vector.</param>
        /// <param name="data">输入，长度须为 16 的倍数。The input, whose length must be a multiple of 16.</param>
        /// <param name="encrypt">true 加密，false 解密。True encrypts, false decrypts.</param>
        /// <returns>结果字节。The result bytes.</returns>
        private static byte[] AesCbc(byte[] key, byte[] iv, byte[] data, bool encrypt)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.None;
                aes.IV = iv;
                using (ICryptoTransform transform = encrypt ? aes.CreateEncryptor() : aes.CreateDecryptor())
                {
                    return transform.TransformFinalBlock(data, 0, data.Length);
                }
            }
        }

        /// <summary>AES-ECB 无填充加解密，用于 /Perms。AES-ECB without padding, used by /Perms.</summary>
        /// <param name="key">32 字节文件密钥。The 32-byte file key.</param>
        /// <param name="data">输入。The input.</param>
        /// <param name="encrypt">true 加密，false 解密。True encrypts, false decrypts.</param>
        /// <returns>结果字节。The result bytes.</returns>
        private static byte[] AesEcb(byte[] key, byte[] data, bool encrypt)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.Mode = CipherMode.ECB;
                aes.Padding = PaddingMode.None;
                using (ICryptoTransform transform = encrypt ? aes.CreateEncryptor() : aes.CreateDecryptor())
                {
                    return transform.TransformFinalBlock(data, 0, data.Length);
                }
            }
        }

        /// <summary>按顺序拼接字节数组。Concatenates byte arrays in order.</summary>
        /// <param name="parts">片段。The segments.</param>
        /// <returns>拼接结果。The concatenated result.</returns>
        private static byte[] Concat(params byte[][] parts)
        {
            int total = 0;
            for (int i = 0; i < parts.Length; i++) total += parts[i].Length;
            byte[] result = new byte[total];
            int offset = 0;
            for (int i = 0; i < parts.Length; i++)
            {
                Buffer.BlockCopy(parts[i], 0, result, offset, parts[i].Length);
                offset += parts[i].Length;
            }
            return result;
        }

        /// <summary>取一段字节。Takes a slice of the bytes.</summary>
        /// <param name="data">输入。The input.</param>
        /// <param name="start">起始位置。The start index.</param>
        /// <param name="length">长度。The length.</param>
        /// <returns>切片。The slice.</returns>
        private static byte[] Slice(byte[] data, int start, int length)
        {
            byte[] result = new byte[length];
            Buffer.BlockCopy(data, start, result, 0, length);
            return result;
        }

        /// <summary>逐字节比较。Compares bytes for equality.</summary>
        /// <param name="a">第一段。The first array.</param>
        /// <param name="b">第二段。The second array.</param>
        /// <returns>相等返回 true。True when equal.</returns>
        private static bool BytesEqual(byte[] a, byte[] b)
        {
            if (a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != b[i]) return false;
            }
            return true;
        }

        /// <summary>PDF 口令填充串（算法 2 的 32 字节常量）。The PDF password padding string (the 32-byte constant of algorithm 2).</summary>
        private static readonly byte[] PasswordPaddingValue =
        {
            0x28, 0xBF, 0x4E, 0x5E, 0x4E, 0x75, 0x8A, 0x41,
            0x64, 0x00, 0x4E, 0x56, 0xFF, 0xFA, 0x01, 0x08,
            0x2E, 0x2E, 0x00, 0xB6, 0xD0, 0x68, 0x3E, 0x80,
            0x2F, 0x0C, 0xA9, 0xFE, 0x64, 0x53, 0x69, 0x7A
        };

        /// <summary>算法 2 的 a 步：口令按字节截断或补齐到 32 位。Step (a) of algorithm 2: truncate or pad the password to 32 bytes.</summary>
        /// <param name="password">口令文本。The password text.</param>
        /// <returns>32 字节。The 32 bytes.</returns>
        private static byte[] PadPassword(string password)
        {
            byte[] text = Encoding.GetEncoding("ISO-8859-1").GetBytes(password ?? string.Empty);
            byte[] result = new byte[32];
            int count = Math.Min(text.Length, 32);
            Buffer.BlockCopy(text, 0, result, 0, count);
            Buffer.BlockCopy(PasswordPaddingValue, 0, result, count, 32 - count);
            return result;
        }

        /// <summary>算法 2：由口令、/O、/P、/ID 与元数据开关推出文件密钥。Algorithm 2: derives the file key from password, /O, /P, /ID and the metadata flag.</summary>
        /// <param name="padded">补齐后的口令。The padded password.</param>
        /// <param name="o">/O 值。The /O value.</param>
        /// <param name="p">/P 值。The /P value.</param>
        /// <param name="id">trailer 的第一个 /ID。The first /ID of the trailer.</param>
        /// <param name="revision">修订号 2 到 4。The revision, 2 to 4.</param>
        /// <param name="keyBits">密钥位数。The key length in bits.</param>
        /// <param name="encryptMetadata">元数据是否加密。Whether metadata is encrypted.</param>
        /// <returns>文件密钥。The file key.</returns>
        private static byte[] ReferenceFileKey(byte[] padded, byte[] o, int p, byte[] id, int revision, int keyBits, bool encryptMetadata)
        {
            byte[] seed = Concat(padded, o, LittleEndianBytes(p), id);
            if (revision >= 4 && !encryptMetadata) seed = Concat(seed, new byte[] { 0xFF, 0xFF, 0xFF, 0xFF });

            int n = Math.Max(5, keyBits / 8);
            byte[] hash = Md5(seed);
            if (revision >= 3)
            {
                for (int i = 0; i < 50; i++) hash = Md5(Slice(hash, 0, n));
            }
            return Slice(hash, 0, n);
        }

        /// <summary>算法 3 的逆：从 /O 还原补齐后的用户口令。The inverse of algorithm 3: recovers the padded user password from /O.</summary>
        /// <param name="o">/O 值。The /O value.</param>
        /// <param name="paddedOwner">补齐后的所有者口令。The padded owner password.</param>
        /// <param name="revision">修订号 2 到 4。The revision, 2 to 4.</param>
        /// <param name="keyBits">密钥位数。The key length in bits.</param>
        /// <returns>补齐后的用户口令。The padded user password.</returns>
        private static byte[] ReferenceOwnerPassword(byte[] o, byte[] paddedOwner, int revision, int keyBits)
        {
            int n = revision == 2 ? 5 : Math.Max(5, keyBits / 8);
            byte[] hash = Md5(paddedOwner);
            if (revision >= 3)
            {
                for (int i = 0; i < 50; i++) hash = Md5(hash);
            }
            byte[] key = Slice(hash, 0, n);

            byte[] value = o;
            if (revision >= 3)
            {
                for (int i = 19; i >= 1; i--) value = Rc4(XorBytes(key, (byte)i), value);
            }
            return Rc4(key, value);
        }

        /// <summary>算法 5：MD5(填充串 + /ID) 之后的 19 轮异或 RC4。Algorithm 5: the 19 XOR-RC4 rounds after MD5(padding + /ID).</summary>
        /// <param name="key">文件密钥。The file key.</param>
        /// <param name="data">MD5(填充串 + /ID) 的摘要。The digest of MD5(padding + /ID).</param>
        /// <returns>32 字节 /U 值。The 32-byte /U value.</returns>
        private static byte[] Rc4Rounds(byte[] key, byte[] data)
        {
            byte[] value = Rc4(key, data);
            for (int i = 1; i <= 19; i++) value = Rc4(XorBytes(key, (byte)i), value);
            return value;
        }

        /// <summary>RC4 对称变换。The RC4 symmetric transform.</summary>
        /// <param name="key">密钥。The key.</param>
        /// <param name="data">输入。The input.</param>
        /// <returns>输出。The output.</returns>
        private static byte[] Rc4(byte[] key, byte[] data)
        {
            int[] state = new int[256];
            for (int i = 0; i < 256; i++) state[i] = i;
            int j = 0;
            for (int i = 0; i < 256; i++)
            {
                j = (j + state[i] + key[i % key.Length]) & 0xFF;
                int temp = state[i];
                state[i] = state[j];
                state[j] = temp;
            }

            byte[] output = new byte[data.Length];
            int a = 0;
            j = 0;
            for (int i = 0; i < data.Length; i++)
            {
                a = (a + 1) & 0xFF;
                j = (j + state[a]) & 0xFF;
                int temp = state[a];
                state[a] = state[j];
                state[j] = temp;
                output[i] = (byte)(data[i] ^ state[(state[a] + state[j]) & 0xFF]);
            }
            return output;
        }

        /// <summary>计算 MD5。Computes MD5.</summary>
        /// <param name="data">输入。The input.</param>
        /// <returns>16 字节摘要。The 16-byte digest.</returns>
        private static byte[] Md5(byte[] data)
        {
            using (MD5 md5 = MD5.Create())
            {
                return md5.ComputeHash(data);
            }
        }

        /// <summary>逐字节异或密钥。XORs every key byte with a value.</summary>
        /// <param name="key">密钥。The key.</param>
        /// <param name="value">异或值。The value.</param>
        /// <returns>新密钥。The new key.</returns>
        private static byte[] XorBytes(byte[] key, byte value)
        {
            byte[] result = new byte[key.Length];
            for (int i = 0; i < key.Length; i++) result[i] = (byte)(key[i] ^ value);
            return result;
        }

        /// <summary>小端 4 字节。The value as four little-endian bytes.</summary>
        /// <param name="value">整数。The integer.</param>
        /// <returns>4 字节。Four bytes.</returns>
        private static byte[] LittleEndianBytes(int value)
        {
            return new byte[]
            {
                (byte)(value & 0xFF),
                (byte)((value >> 8) & 0xFF),
                (byte)((value >> 16) & 0xFF),
                (byte)((value >> 24) & 0xFF)
            };
        }

        /// <summary>读 trailer 中第一个 /ID 串。Reads the first /ID string from the trailer.</summary>
        /// <param name="file">已打开的文档。The opened document.</param>
        /// <returns>标识字节。The identifier bytes.</returns>
        private static byte[] FirstDocumentId(PdfFile file)
        {
            PdfArray ids = file.Resolve(file.Trailer.Get("ID")) as PdfArray;
            Require(ids != null && ids.Count > 0, "an encrypted document must carry a trailer /ID array");
            PdfString text = file.Resolve(ids[0]) as PdfString;
            Require(text != null, "the first /ID entry must be a string");
            return text.Bytes;
        }
    }
}
