using System;
using System.IO;
using System.Text;

namespace Wisdom.Pdf
{
    /// <summary>
    /// PDF 标准安全处理器：支持 RC4-40(R2)、RC4-128(R3)、AES-128(R4) 与 AES-256(R5/R6)。
    /// The PDF standard security handler: supports RC4-40 (R2), RC4-128 (R3), AES-128 (R4) and
    /// AES-256 (R5/R6).
    /// </summary>
    /// <remarks>
    /// 读取端按 <c>/V</c>、<c>/R</c> 与密码字典中的 <c>/CryptFilter</c> 选择算法，先以用户口令校验 <c>/U</c>，
    /// 失败后以所有者口令还原用户口令再校验；写出端按 <see cref="PdfEncryptionOptions"/> 生成 O/U/OE/UE/Perms。
    /// 元数据流在 <c>/EncryptMetadata</c> 为 false 时保持明文。
    /// The reader selects the algorithm from <c>/V</c>, <c>/R</c> and the crypt filters, checking <c>/U</c>
    /// with the user password first and, failing that, recovering the user password from the owner password.
    /// The writer derives O/U/OE/UE/Perms from <see cref="PdfEncryptionOptions"/>. Metadata streams stay in
    /// plain text when <c>/EncryptMetadata</c> is false.
    /// </remarks>
    internal sealed class PdfStandardSecurityHandler
    {
        /// <summary>PDF 口令填充串（ISO 32000-1，算法 2 的 32 字节常量）。The PDF password padding string (the 32-byte constant of Algorithm 2).</summary>
        private static readonly byte[] PasswordPadding =
        {
            0x28, 0xBF, 0x4E, 0x5E, 0x4E, 0x75, 0x8A, 0x41,
            0x64, 0x00, 0x4E, 0x56, 0xFF, 0xFA, 0x01, 0x08,
            0x2E, 0x2E, 0x00, 0xB6, 0xD0, 0x68, 0x3E, 0x80,
            0x2F, 0x0C, 0xA9, 0xFE, 0x64, 0x53, 0x69, 0x7A
        };

        /// <summary>AES 对象密钥派生时追加的 "sAlT" 标记。The "sAlT" marker appended when deriving AES object keys.</summary>
        private static readonly byte[] AesSalt = { 0x73, 0x41, 0x6C, 0x54 };

        /// <summary>加密字典。The encryption dictionary.</summary>
        private readonly PdfDictionary _dictionary;
        /// <summary>文件加密密钥。The file encryption key.</summary>
        private readonly byte[] _key;
        /// <summary>生效权限；所有者口令打开时为全部权限。The effective permissions; all permissions when opened with the owner password.</summary>
        private readonly PdfPermissions _permissions;
        /// <summary>是否加密元数据流。Whether metadata streams are encrypted.</summary>
        private readonly bool _encryptMetadata;
        /// <summary>安全处理器修订号。The security handler revision.</summary>
        private readonly int _revision;
        /// <summary>文件密钥位数。The file key length in bits.</summary>
        private readonly int _keyBits;
        /// <summary>是否使用 AES-256（R5/R6）。Whether AES-256 (R5/R6) is used.</summary>
        private readonly bool _aes256;
        /// <summary>字符串是否加密（/StrF 非 Identity）。Whether strings are encrypted (/StrF is not Identity).</summary>
        private readonly bool _encryptStrings;
        /// <summary>流是否加密（/StmF 非 Identity）。Whether streams are encrypted (/StmF is not Identity).</summary>
        private readonly bool _encryptStreams;
        /// <summary>字符串是否使用 AES（否则为 RC4）。Whether strings use AES (otherwise RC4).</summary>
        private readonly bool _stringAes;
        /// <summary>流是否使用 AES（否则为 RC4）。Whether streams use AES (otherwise RC4).</summary>
        private readonly bool _streamAes;
        /// <summary>是否以所有者口令通过校验。Whether the owner password passed authentication.</summary>
        private readonly bool _owner;

        /// <summary>
        /// 由 <see cref="Open"/> 或 <see cref="Create"/> 创建处理器。
        /// Creates a handler through <see cref="Open"/> or <see cref="Create"/>.
        /// </summary>
        private PdfStandardSecurityHandler(PdfDictionary dictionary, byte[] key,
            PdfPermissions permissions, bool encryptMetadata, int revision, int keyBits,
            bool aes256, bool encryptStrings, bool encryptStreams, bool stringAes, bool streamAes, bool owner)
        {
            _dictionary = dictionary;
            _key = key;
            _permissions = permissions;
            _encryptMetadata = encryptMetadata;
            _revision = revision;
            _keyBits = keyBits;
            _aes256 = aes256;
            _encryptStrings = encryptStrings;
            _encryptStreams = encryptStreams;
            _stringAes = stringAes;
            _streamAes = streamAes;
            _owner = owner;
        }

        /// <summary>加密字典。The encryption dictionary.</summary>
        public PdfDictionary EncryptionDictionary { get { return _dictionary; } }

        /// <summary>生效权限，所有者口令打开时为全部权限。The effective permissions; all permissions when opened as the owner.</summary>
        public PdfPermissions Permissions { get { return _permissions; } }

        /// <summary>是否加密元数据流。Whether metadata streams are encrypted.</summary>
        public bool EncryptMetadata { get { return _encryptMetadata; } }

        /// <summary>是否以所有者口令通过校验。Whether the owner password was used.</summary>
        public bool IsOwner { get { return _owner; } }

        /// <summary>修订号（2–6）。The revision number (2-6).</summary>
        public int Revision { get { return _revision; } }

        /// <summary>文件密钥位数。The file key length in bits.</summary>
        public int KeyBits { get { return _keyBits; } }

        /// <summary>
        /// 算法名称，用于诊断与界面显示。
        /// The algorithm name, used for diagnostics and display.
        /// </summary>
        public string AlgorithmName
        {
            get
            {
                if (_revision >= 5) return "AES-256 (R" + _revision.ToString() + ")";
                if (_stringAes || _streamAes) return "AES-128 (R4)";
                return _keyBits <= 40 ? "RC4-40 (R2)" : "RC4-128 (R" + _revision.ToString() + ")";
            }
        }

        /// <summary>
        /// 判断一个流字典是否声明为元数据流。
        /// Determines whether a stream dictionary is declared as a metadata stream.
        /// </summary>
        /// <param name="dictionary">流字典。The stream dictionary.</param>
        /// <returns>是元数据返回 true。True for metadata streams.</returns>
        public static bool IsMetadataStream(PdfDictionary dictionary)
        {
            if (dictionary == null) return false;
            PdfName type = dictionary.Get("Type") as PdfName;
            return type != null && type.Value == "Metadata";
        }

        /// <summary>
        /// 打开受口令保护的文档：校验口令并建立文件密钥。
        /// Opens a password-protected document: validates the password and establishes the file key.
        /// </summary>
        /// <param name="dictionary">加密字典。The encryption dictionary.</param>
        /// <param name="documentId">trailer 的第一个 /ID 串。The first /ID string of the trailer.</param>
        /// <param name="password">用户或所有者口令，null 视为空口令。The user or owner password; null means empty.</param>
        /// <returns>已认证的处理器。The authenticated handler.</returns>
        /// <exception cref="PdfPasswordException">口令错误。The password is incorrect.</exception>
        /// <exception cref="NotSupportedException">安全处理器或算法不受支持。The handler or algorithm is unsupported.</exception>
        public static PdfStandardSecurityHandler Open(PdfDictionary dictionary, byte[] documentId, string password)
        {
            if (dictionary == null) throw new ArgumentNullException("dictionary");

            string filter = GetName(dictionary, "Filter");
            if (!string.IsNullOrEmpty(filter) && filter != "Standard")
            {
                throw new NotSupportedException("Unsupported PDF security handler: " + filter);
            }

            int version = GetInt(dictionary, "V", 0);
            int revision = GetInt(dictionary, "R", 2);
            if (revision < 2 || revision > 6)
            {
                throw new NotSupportedException("Unsupported PDF security revision: R" + revision.ToString());
            }
            if (version == 0) version = revision >= 5 ? 5 : (revision == 2 ? 1 : (revision >= 4 ? 4 : 2));

            int keyBits = GetInt(dictionary, "Length", 0);
            if (version >= 5 || revision >= 5) keyBits = 256;
            else if (keyBits < 40 || keyBits > 128) keyBits = revision == 2 ? 40 : 128;
            else keyBits = ((keyBits + 7) / 8) * 8;

            bool encryptMetadata = !IsFalse(dictionary.Get("EncryptMetadata"));
            int permissionsValue = GetInt(dictionary, "P", 0);
            PdfPermissions permissions = (PdfPermissions)(permissionsValue & (int)PdfPermissions.All);

            if (documentId == null) documentId = new byte[0];
            string candidate = password ?? string.Empty;

            // 先定算法再验口令：/CF 无法解析时继续做密钥派生只会把算法混淆掩盖在口令错误之下。
            // The algorithm is settled before the password is checked: deriving a key when /CF cannot be
            // resolved would hide an algorithm confusion behind a password error.
            bool encryptStrings = true;
            bool encryptStreams = true;
            bool stringAes = version >= 5;
            bool streamAes = version >= 5;

            if (version >= 4)
            {
                string streamFilter = GetName(dictionary, "StmF");
                string stringFilter = GetName(dictionary, "StrF");
                string streamMethod = GetCryptFilterMethod(dictionary, streamFilter);
                string stringMethod = GetCryptFilterMethod(dictionary, stringFilter);
                encryptStreams = streamMethod != null;
                encryptStrings = stringMethod != null;
                streamAes = streamMethod == "AESV2" || streamMethod == "AESV3";
                stringAes = stringMethod == "AESV2" || stringMethod == "AESV3";
                if (version >= 5)
                {
                    if (encryptStreams) streamAes = true;
                    if (encryptStrings) stringAes = true;
                }
            }

            byte[] key;
            bool owner = false;

            if (revision >= 5)
            {
                key = OpenAes256(dictionary, candidate, revision, ref owner);
            }
            else
            {
                byte[] o = GetBytes(dictionary, "O");
                byte[] u = GetBytes(dictionary, "U");
                if (u.Length < 16) throw new InvalidDataException("The PDF encryption dictionary has an invalid /U value.");

                key = ComputeFileKey(PadPassword(candidate, 32), o, permissionsValue, documentId, revision, keyBits, encryptMetadata);
                byte[] expected = ComputeUserValue(key, documentId, revision);
                bool user = CompareUser(expected, u, revision);

                if (!user)
                {
                    byte[] ownerKey = ComputeOwnerKey(candidate, revision, keyBits);
                    byte[] paddedUser = DecryptOwnerValue(o, ownerKey, revision);
                    byte[] ownerFileKey = ComputeFileKey(paddedUser, o, permissionsValue, documentId, revision, keyBits, encryptMetadata);
                    byte[] ownerExpected = ComputeUserValue(ownerFileKey, documentId, revision);
                    if (CompareUser(ownerExpected, u, revision))
                    {
                        key = ownerFileKey;
                        owner = true;
                    }
                }

                if (!user && !owner)
                {
                    throw new PdfPasswordException("The password is incorrect.");
                }
            }


            if (owner) permissions = PdfPermissions.All;
            return new PdfStandardSecurityHandler(dictionary, key, permissions, encryptMetadata,
                revision, keyBits, version >= 5, encryptStrings, encryptStreams, stringAes, streamAes, owner);
        }

        /// <summary>
        /// 按选项创建加密字典与文件密钥，供写出端使用。
        /// Creates an encryption dictionary and file key from the options, for the writer.
        /// </summary>
        /// <param name="options">加密选项。The encryption options.</param>
        /// <param name="documentId">trailer 的第一个 /ID 串，R2–R4 参与密钥派生。The first /ID string, used by R2-R4 key derivation.</param>
        /// <returns>新建的处理器。The new handler.</returns>
        public static PdfStandardSecurityHandler Create(PdfEncryptionOptions options, byte[] documentId)
        {
            if (options == null) options = new PdfEncryptionOptions();
            if (documentId == null) documentId = new byte[0];

            string user = options.UserPassword ?? string.Empty;
            string owner = string.IsNullOrEmpty(options.OwnerPassword) ? user : options.OwnerPassword;
            int permissionsValue = ComputePermissionsValue(options.Permissions);
            bool encryptMetadata = options.EncryptMetadata;

            int version;
            int revision;
            int keyBits;
            bool aes;
            bool aes256;
            switch (options.Algorithm)
            {
                case PdfEncryptionAlgorithm.Rc4_40:
                    version = 1;
                    revision = 2;
                    keyBits = 40;
                    aes = false;
                    aes256 = false;
                    break;
                case PdfEncryptionAlgorithm.Rc4_128:
                    version = 2;
                    revision = 3;
                    keyBits = 128;
                    aes = false;
                    aes256 = false;
                    break;
                case PdfEncryptionAlgorithm.Aes128:
                    version = 4;
                    revision = 4;
                    keyBits = 128;
                    aes = true;
                    aes256 = false;
                    break;
                default:
                    version = 5;
                    revision = 6;
                    keyBits = 256;
                    aes = true;
                    aes256 = true;
                    break;
            }

            PdfDictionary dictionary = new PdfDictionary();
            dictionary.SetName("Filter", "Standard");
            dictionary.SetNumber("V", version);
            dictionary.SetNumber("R", revision);
            dictionary.SetNumber("Length", keyBits);
            dictionary.SetNumber("P", permissionsValue);
            dictionary.SetBoolean("EncryptMetadata", encryptMetadata);

            byte[] key;
            byte[] o;
            byte[] u;
            byte[] oe = null;
            byte[] ue = null;
            byte[] perms = null;

            if (aes256)
            {
                key = PdfCrypto.Random(32);
                byte[] userPassword = EncodeUtf8Password(user);
                byte[] ownerPassword = EncodeUtf8Password(owner);

                byte[] userValidationSalt = PdfCrypto.Random(8);
                byte[] userKeySalt = PdfCrypto.Random(8);
                byte[] userHash = HashPassword(userPassword, userValidationSalt, null, revision);
                u = Concat(Truncate(userHash, 32), userValidationSalt, userKeySalt);
                byte[] userKey = HashPassword(userPassword, userKeySalt, null, revision);
                ue = PdfCrypto.AesEncryptNoPadding(Truncate(userKey, 32), new byte[16], key);

                byte[] ownerValidationSalt = PdfCrypto.Random(8);
                byte[] ownerKeySalt = PdfCrypto.Random(8);
                byte[] userValue48 = Truncate(u, 48);
                byte[] ownerHash = HashPassword(ownerPassword, ownerValidationSalt, userValue48, revision);
                o = Concat(Truncate(ownerHash, 32), ownerValidationSalt, ownerKeySalt);
                byte[] ownerKey = HashPassword(ownerPassword, ownerKeySalt, userValue48, revision);
                oe = PdfCrypto.AesEncryptNoPadding(Truncate(ownerKey, 32), new byte[16], key);

                perms = PdfCrypto.AesEcbEncrypt(key, BuildPermissionsBlock(permissionsValue, encryptMetadata));

                dictionary.Set("O", new PdfString(o, true));
                dictionary.Set("U", new PdfString(u, true));
                dictionary.Set("OE", new PdfString(oe, true));
                dictionary.Set("UE", new PdfString(ue, true));
                dictionary.Set("Perms", new PdfString(perms, true));
                SetCryptFilter(dictionary, "AESV3", 32);
            }
            else
            {
                o = ComputeOwnerValue(owner, user, revision, keyBits);
                key = ComputeFileKey(PadPassword(user, 32), o, permissionsValue, documentId, revision, keyBits, encryptMetadata);
                u = ComputeUserValue(key, documentId, revision);

                dictionary.Set("O", new PdfString(o, true));
                dictionary.Set("U", new PdfString(u, true));
                if (aes) SetCryptFilter(dictionary, "AESV2", 16);
            }

            return new PdfStandardSecurityHandler(dictionary, key,
                options.Permissions & PdfPermissions.All, encryptMetadata, revision, keyBits, aes256,
                true, true, aes, aes, false);
        }

        /// <summary>
        /// 将明文字符串加密为密文字节。
        /// Encrypts a plaintext string into ciphertext bytes.
        /// </summary>
        /// <param name="data">明文。The plaintext.</param>
        /// <param name="number">所属对象号。The owning object number.</param>
        /// <param name="generation">所属代次。The owning generation.</param>
        /// <returns>密文。The ciphertext.</returns>
        public byte[] EncryptString(byte[] data, int number, int generation)
        {
            if (!_encryptStrings) return data;
            return Encrypt(data, number, generation, _stringAes);
        }

        /// <summary>
        /// 将明文流数据加密为密文字节。
        /// Encrypts plaintext stream data into ciphertext bytes.
        /// </summary>
        /// <param name="data">明文。The plaintext.</param>
        /// <param name="number">所属对象号。The owning object number.</param>
        /// <param name="generation">所属代次。The owning generation.</param>
        /// <returns>密文。The ciphertext.</returns>
        public byte[] EncryptStream(byte[] data, int number, int generation)
        {
            if (!_encryptStreams) return data;
            return Encrypt(data, number, generation, _streamAes);
        }

        /// <summary>
        /// 将密文字符串解密为明文。
        /// Decrypts a ciphertext string into plaintext.
        /// </summary>
        /// <param name="data">密文。The ciphertext.</param>
        /// <param name="number">所属对象号。The owning object number.</param>
        /// <param name="generation">所属代次。The owning generation.</param>
        /// <returns>明文。The plaintext.</returns>
        public byte[] DecryptString(byte[] data, int number, int generation)
        {
            if (!_encryptStrings) return data;
            return Decrypt(data, number, generation, _stringAes);
        }

        /// <summary>
        /// 将密文流数据解密为明文。
        /// Decrypts ciphertext stream data into plaintext.
        /// </summary>
        /// <param name="data">密文。The ciphertext.</param>
        /// <param name="number">所属对象号。The owning object number.</param>
        /// <param name="generation">所属代次。The owning generation.</param>
        /// <returns>明文。The plaintext.</returns>
        public byte[] DecryptStream(byte[] data, int number, int generation)
        {
            if (!_encryptStreams) return data;
            return Decrypt(data, number, generation, _streamAes);
        }

        /// <summary>
        /// 按算法加密一段数据；空数据保持为空。
        /// Encrypts one piece of data with the configured algorithm; empty data stays empty.
        /// </summary>
        private byte[] Encrypt(byte[] data, int number, int generation, bool aes)
        {
            if (data == null || data.Length == 0) return new byte[0];
            if (aes)
            {
                byte[] key = _aes256 ? _key : ComputeObjectKey(number, generation, true);
                return PdfCrypto.AesEncrypt(key, data);
            }
            return PdfCrypto.Rc4(ComputeObjectKey(number, generation, false), data);
        }

        /// <summary>
        /// 按算法解密一段数据；空数据保持为空。
        /// Decrypts one piece of data with the configured algorithm; empty data stays empty.
        /// </summary>
        private byte[] Decrypt(byte[] data, int number, int generation, bool aes)
        {
            if (data == null || data.Length == 0) return new byte[0];
            if (aes)
            {
                byte[] key = _aes256 ? _key : ComputeObjectKey(number, generation, true);
                return PdfCrypto.AesDecrypt(key, data);
            }
            return PdfCrypto.Rc4(ComputeObjectKey(number, generation, false), data);
        }

        /// <summary>
        /// 由文件密钥派生单个对象使用的密钥：MD5(文件密钥 + 对象号低 3 字节 + 代次低 2 字节 [+ sAlT])，取前 n+5 字节。
        /// Derives the per-object key from the file key: MD5(file key + low 3 object bytes + low 2 generation
        /// bytes [+ sAlT]), truncated to n+5 bytes.
        /// </summary>
        /// <param name="number">对象号。The object number.</param>
        /// <param name="generation">代次。The generation.</param>
        /// <param name="aes">是否为 AES（需要追加 sAlT）。Whether AES is used (which appends sAlT).</param>
        /// <returns>对象密钥。The object key.</returns>
        private byte[] ComputeObjectKey(int number, int generation, bool aes)
        {
            byte[] objectBytes =
            {
                (byte)(number & 0xFF),
                (byte)((number >> 8) & 0xFF),
                (byte)((number >> 16) & 0xFF),
                (byte)(generation & 0xFF),
                (byte)((generation >> 8) & 0xFF)
            };
            byte[] hash = aes
                ? PdfCrypto.Md5(_key, objectBytes, AesSalt)
                : PdfCrypto.Md5(_key, objectBytes);
            int length = Math.Min(16, _key.Length + 5);
            return Truncate(hash, length);
        }

        /// <summary>
        /// R2–R4：由已填充的用户口令推导文件密钥（算法 2）。
        /// R2-R4: derives the file key from the padded user password (Algorithm 2).
        /// </summary>
        private static byte[] ComputeFileKey(byte[] paddedPassword, byte[] o, int permissions, byte[] documentId,
            int revision, int keyBits, bool encryptMetadata)
        {
            int length = Math.Max(5, keyBits / 8);
            byte[] documentBytes = documentId ?? new byte[0];
            byte[] hash;
            if (revision >= 4 && !encryptMetadata)
            {
                hash = PdfCrypto.Md5(paddedPassword, o, PdfCrypto.LittleEndian(permissions), documentBytes,
                    new byte[] { 0xFF, 0xFF, 0xFF, 0xFF });
            }
            else
            {
                hash = PdfCrypto.Md5(paddedPassword, o, PdfCrypto.LittleEndian(permissions), documentBytes);
            }

            byte[] result = Truncate(hash, length);
            if (revision >= 3)
            {
                for (int i = 0; i < 50; i++)
                {
                    result = Truncate(PdfCrypto.Md5(result), length);
                }
            }
            return result;
        }

        /// <summary>
        /// R2–R4：计算 /U（算法 4 与 5）。
        /// R2-R4: computes /U (Algorithms 4 and 5).
        /// </summary>
        private static byte[] ComputeUserValue(byte[] key, byte[] documentId, int revision)
        {
            if (revision == 2)
            {
                return PdfCrypto.Rc4(key, PasswordPadding);
            }

            byte[] hash = PdfCrypto.Md5(PasswordPadding, documentId ?? new byte[0]);
            byte[] value = PdfCrypto.Rc4(key, hash);
            for (int i = 1; i <= 19; i++)
            {
                value = PdfCrypto.Rc4(XorKey(key, (byte)i), value);
            }

            byte[] result = new byte[32];
            Buffer.BlockCopy(value, 0, result, 0, 16);
            return result;
        }

        /// <summary>
        /// R2–R4：比较 /U，R3 起只比较前 16 字节。
        /// R2-R4: compares /U, only the first 16 bytes from R3 on.
        /// </summary>
        private static bool CompareUser(byte[] expected, byte[] actual, int revision)
        {
            if (actual == null || actual.Length < 16) return false;
            int length = revision == 2 ? Math.Min(32, actual.Length) : 16;
            byte[] left = Truncate(expected, length);
            byte[] right = Truncate(actual, length);
            return PdfCrypto.BytesEqual(left, right);
        }

        /// <summary>
        /// R2–R4：计算 /O（算法 3），所有者口令为空时使用用户口令。
        /// R2-R4: computes /O (Algorithm 3); the user password is used when the owner password is empty.
        /// </summary>
        private static byte[] ComputeOwnerValue(string ownerPassword, string userPassword, int revision, int keyBits)
        {
            string password = string.IsNullOrEmpty(ownerPassword) ? (userPassword ?? string.Empty) : ownerPassword;
            int length = Math.Max(5, keyBits / 8);
            byte[] hash = PdfCrypto.Md5(PadPassword(password, 32));
            byte[] key = Truncate(hash, length);
            if (revision >= 3)
            {
                for (int i = 0; i < 50; i++)
                {
                    key = Truncate(PdfCrypto.Md5(key), length);
                }
            }

            byte[] value = PdfCrypto.Rc4(key, PadPassword(userPassword ?? string.Empty, 32));
            if (revision >= 3)
            {
                for (int i = 1; i <= 19; i++)
                {
                    value = PdfCrypto.Rc4(XorKey(key, (byte)i), value);
                }
            }
            return value;
        }

        /// <summary>
        /// R2–R4：由所有者口令推导 RC4 密钥（算法 3 的前三步）。
        /// R2-R4: derives the RC4 key from the owner password (the first three steps of Algorithm 3).
        /// </summary>
        private static byte[] ComputeOwnerKey(string ownerPassword, int revision, int keyBits)
        {
            int length = Math.Max(5, keyBits / 8);
            byte[] hash = PdfCrypto.Md5(PadPassword(ownerPassword ?? string.Empty, 32));
            byte[] key = Truncate(hash, length);
            if (revision >= 3)
            {
                for (int i = 0; i < 50; i++)
                {
                    key = Truncate(PdfCrypto.Md5(key), length);
                }
            }
            return key;
        }

        /// <summary>
        /// R2–R4：用所有者密钥解密 /O，还原填充后的用户口令。
        /// R2-R4: decrypts /O with the owner key, recovering the padded user password.
        /// </summary>
        private static byte[] DecryptOwnerValue(byte[] o, byte[] ownerKey, int revision)
        {
            byte[] value = o ?? new byte[0];
            if (revision == 2) return PdfCrypto.Rc4(ownerKey, value);

            for (int i = 19; i >= 1; i--)
            {
                value = PdfCrypto.Rc4(XorKey(ownerKey, (byte)i), value);
            }
            return PdfCrypto.Rc4(ownerKey, value);
        }

        /// <summary>
        /// R5/R6：校验口令并解出 32 字节文件密钥。
        /// R5/R6: validates the password and recovers the 32-byte file key.
        /// </summary>
        private static byte[] OpenAes256(PdfDictionary dictionary, string password, int revision, ref bool owner)
        {
            byte[] u = GetBytes(dictionary, "U");
            byte[] o = GetBytes(dictionary, "O");
            byte[] ue = GetBytes(dictionary, "UE");
            byte[] oe = GetBytes(dictionary, "OE");
            if (u.Length < 48 || o.Length < 48 || ue.Length < 32 || oe.Length < 32)
            {
                throw new InvalidDataException("The PDF encryption dictionary is incomplete.");
            }

            byte[] encoded = EncodeUtf8Password(password);
            byte[] userHash = HashPassword(encoded, Slice(u, 32, 8), null, revision);
            if (CompareHash(userHash, u))
            {
                byte[] userKey = HashPassword(encoded, Slice(u, 40, 8), null, revision);
                return PdfCrypto.AesDecryptNoPadding(Truncate(userKey, 32), new byte[16], Slice(ue, 0, 32));
            }

            byte[] userValue48 = Slice(u, 0, 48);
            byte[] ownerHash = HashPassword(encoded, Slice(o, 32, 8), userValue48, revision);
            if (CompareHash(ownerHash, o))
            {
                owner = true;
                byte[] ownerKey = HashPassword(encoded, Slice(o, 40, 8), userValue48, revision);
                return PdfCrypto.AesDecryptNoPadding(Truncate(ownerKey, 32), new byte[16], Slice(oe, 0, 32));
            }

            throw new PdfPasswordException("The password is incorrect.");
        }

        /// <summary>
        /// R5/R6 的口令散列：R5 为 SHA-256，R6 使用算法 2.B 的迭代加固。
        /// The R5/R6 password hash: SHA-256 for R5, the iterated Algorithm 2.B for R6.
        /// </summary>
        private static byte[] HashPassword(byte[] password, byte[] salt, byte[] userKey, int revision)
        {
            if (revision <= 5) return PdfCrypto.Sha256(password, salt, userKey);
            return HashRevision6(password, salt, userKey);
        }

        /// <summary>
        /// R6 的算法 2.B：以 SHA-256 起算，按 AES 密文前 16 字节之和中选 SHA-256/384/512，至少 64 轮。
        /// R6 Algorithm 2.B: starts from SHA-256, picks SHA-256/384/512 from the modular sum of the first 16
        /// bytes of the AES ciphertext, and runs at least 64 rounds.
        /// </summary>
        private static byte[] HashRevision6(byte[] password, byte[] salt, byte[] userKey)
        {
            byte[] key = PdfCrypto.Sha256(password, salt, userKey);
            int fixedLength = password.Length + (userKey == null ? 0 : userKey.Length);
            int rounds = 0;

            while (true)
            {
                int blockLength = (fixedLength + key.Length) * 64;
                byte[] block = new byte[blockLength];
                int offset = 0;
                for (int repeat = 0; repeat < 64; repeat++)
                {
                    Buffer.BlockCopy(password, 0, block, offset, password.Length);
                    offset += password.Length;
                    Buffer.BlockCopy(key, 0, block, offset, key.Length);
                    offset += key.Length;
                    if (userKey != null && userKey.Length > 0)
                    {
                        Buffer.BlockCopy(userKey, 0, block, offset, userKey.Length);
                        offset += userKey.Length;
                    }
                }

                byte[] aesKey = Slice(key, 0, 16);
                byte[] iv = Slice(key, 16, 16);
                byte[] encrypted = PdfCrypto.AesEncryptNoPadding(aesKey, iv, block);

                int sum = 0;
                for (int i = 0; i < 16; i++) sum += encrypted[i] & 0xFF;
                switch (sum % 3)
                {
                    case 0:
                        key = PdfCrypto.Sha256(encrypted);
                        break;
                    case 1:
                        key = PdfCrypto.Sha384(encrypted);
                        break;
                    default:
                        key = PdfCrypto.Sha512(encrypted);
                        break;
                }

                rounds++;
                if (rounds >= 64 && (encrypted[encrypted.Length - 1] & 0xFF) <= rounds - 32) break;
            }

            return Truncate(key, 32);
        }

        /// <summary>
        /// 比较 32 字节口令散列。
        /// Compares a 32-byte password hash.
        /// </summary>
        private static bool CompareHash(byte[] hash, byte[] value)
        {
            return PdfCrypto.BytesEqual(Truncate(hash, 32), Slice(value, 0, 32));
        }

        /// <summary>
        /// 构造 R5/R6 的 16 字节 /Perms 块：权限、0xFFFFFFFF、元数据标志、"adb" 与随机尾部。
        /// Builds the 16-byte R5/R6 /Perms block: permissions, 0xFFFFFFFF, the metadata flag, "adb" and a
        /// random tail.
        /// </summary>
        private static byte[] BuildPermissionsBlock(int permissions, bool encryptMetadata)
        {
            byte[] block = new byte[16];
            Buffer.BlockCopy(PdfCrypto.LittleEndian(permissions), 0, block, 0, 4);
            block[4] = 0xFF;
            block[5] = 0xFF;
            block[6] = 0xFF;
            block[7] = 0xFF;
            block[8] = (byte)(encryptMetadata ? 'T' : 'F');
            block[9] = (byte)'a';
            block[10] = (byte)'d';
            block[11] = (byte)'b';
            byte[] tail = PdfCrypto.Random(4);
            Buffer.BlockCopy(tail, 0, block, 12, 4);
            return block;
        }

        /// <summary>
        /// 把权限位转换为 /P：低两位清零，第 13 位起全部置 1。
        /// Converts permission flags into /P: the two low bits are cleared and every bit from 13 on is set.
        /// </summary>
        private static int ComputePermissionsValue(PdfPermissions permissions)
        {
            return (int)(((uint)permissions & 0x00000F3Cu) | 0xFFFFF0C0u);
        }

        /// <summary>
        /// R2–R4 口令填充：按 PDFDocEncoding 取低字节，截断到 32 字节后用固定串补齐。
        /// R2-R4 password padding: takes low bytes (PDFDocEncoding), truncates to 32 bytes and pads with the
        /// fixed string.
        /// </summary>
        private static byte[] PadPassword(string password, int length)
        {
            if (password == null) password = string.Empty;
            byte[] result = new byte[32];
            int count = Math.Min(Math.Min(password.Length, length), 32);
            for (int i = 0; i < count; i++)
            {
                char c = password[i];
                result[i] = c <= 0xFF ? (byte)c : (byte)'?';
            }
            for (int i = count; i < 32; i++)
            {
                result[i] = PasswordPadding[i - count];
            }
            return result;
        }

        /// <summary>
        /// R5/R6 口令编码：UTF-8，最多 127 字节。
        /// R5/R6 password encoding: UTF-8, at most 127 bytes.
        /// </summary>
        private static byte[] EncodeUtf8Password(string password)
        {
            if (string.IsNullOrEmpty(password)) return new byte[0];
            byte[] bytes = Encoding.UTF8.GetBytes(password);
            return Truncate(bytes, 127);
        }

        /// <summary>
        /// 按位异或密钥的每个字节，用于 R3 起的 19 轮 RC4。
        /// XORs every key byte with a value, used by the 19 RC4 rounds from R3 on.
        /// </summary>
        private static byte[] XorKey(byte[] key, byte value)
        {
            byte[] result = new byte[key.Length];
            for (int i = 0; i < key.Length; i++) result[i] = (byte)(key[i] ^ value);
            return result;
        }

        /// <summary>
        /// 写出 V4/V5 的 crypt filter 字典。
        /// Writes the V4/V5 crypt filter dictionary.
        /// </summary>
        private static void SetCryptFilter(PdfDictionary dictionary, string method, int lengthBytes)
        {
            PdfDictionary standard = new PdfDictionary();
            standard.SetName("Type", "CryptFilter");
            standard.SetName("CFM", method);
            standard.SetNumber("Length", lengthBytes);
            standard.SetName("AuthEvent", "DocOpen");
            PdfDictionary filters = new PdfDictionary();
            filters.Set("StdCF", standard);
            dictionary.Set("CF", filters);
            dictionary.SetName("StmF", "StdCF");
            dictionary.SetName("StrF", "StdCF");
        }

        /// <summary>
        /// 读取 crypt filter 的加密方法；Identity 或缺失时返回 null。
        /// Reads the crypt filter method; returns null for Identity or when absent.
        /// </summary>
        // /StmF 或 /StrF 点名了某个密码过滤器却在 /CF 里找不到、或 CFM 认不出来时，
        // 静默按不加密处理会把密文当成明文交给上层；这种情况必须报错。
        // When /StmF or /StrF names a crypt filter that cannot be resolved in /CF, or whose /CFM is not
        // recognised, silently treating the file as unencrypted hands ciphertext to the caller as if it
        // were plaintext; that case must raise instead.
        private static string GetCryptFilterMethod(PdfDictionary dictionary, string filterName)
        {
            if (string.IsNullOrEmpty(filterName) || filterName == "Identity") return null;
            // 规范把 /CF 列为 /StmF 或 /StrF 存在时的必填项，缺席或未定义同样无法判定算法。
            // The specification makes /CF a required entry whenever /StmF or /StrF is present; a missing /CF
            // or an undefined name leaves the algorithm undetermined either way.
            if (!string.IsNullOrEmpty(filterName) && filterName != "Identity")
            {
                PdfDictionary namedFilters = dictionary.Get("CF") as PdfDictionary;
                if (namedFilters == null || namedFilters.Get(filterName) == null)
                {
                    throw new InvalidDataException("The encryption dictionary names a crypt filter that /CF does not define: " + filterName);
                }
            }
            PdfDictionary filters = dictionary.Get("CF") as PdfDictionary;
            if (filters == null) return null;
            PdfDictionary filter = filters.Get(filterName) as PdfDictionary;
            if (filter == null) return null;
            string method = GetName(filter, "CFM");
            if (string.IsNullOrEmpty(method) || method == "None") return null;
            if (method == "V2" || method == "AESV2" || method == "AESV3") return method;
            throw new InvalidDataException("Unsupported crypt filter method: " + method);
        }

        /// <summary>读取字典中的名称值。Reads a name value from the dictionary.</summary>
        private static string GetName(PdfDictionary dictionary, string key)
        {
            if (dictionary == null) return null;
            PdfName name = dictionary.Get(key) as PdfName;
            return name == null ? null : name.Value;
        }

        /// <summary>读取字典中的整数值。Reads an integer value from the dictionary.</summary>
        private static int GetInt(PdfDictionary dictionary, string key, int fallback)
        {
            if (dictionary == null) return fallback;
            PdfNumber number = dictionary.Get(key) as PdfNumber;
            if (number == null) return fallback;
            // /P 规范上是有符号 32 位整数，但部分写出器会写成 4294967292 这样的无符号值；
            // 取整后截断低 32 位，才能还原写出方实际使用的字节序列。
            // /P is a signed 32-bit integer per spec, yet some writers emit it unsigned as 4294967292;
            // rounding and truncating to the low 32 bits reproduces the byte sequence the writer used.
            double value = number.Value;
            if (double.IsNaN(value) || double.IsInfinity(value)) return fallback;
            return unchecked((int)(long)Math.Round(value));
        }

        /// <summary>读取字典中的字符串字节。Reads string bytes from the dictionary.</summary>
        private static byte[] GetBytes(PdfDictionary dictionary, string key)
        {
            if (dictionary == null) return new byte[0];
            PdfString text = dictionary.Get(key) as PdfString;
            return text == null ? new byte[0] : text.Bytes;
        }

        /// <summary>判断对象是否为布尔 false。Determines whether an object is boolean false.</summary>
        private static bool IsFalse(PdfObject value)
        {
            PdfBoolean boolean = value as PdfBoolean;
            return boolean != null && !boolean.Value;
        }

        /// <summary>取数组前若干字节。Takes the leading bytes of an array.</summary>
        private static byte[] Truncate(byte[] value, int length)
        {
            if (value == null) return new byte[0];
            if (value.Length <= length) return (byte[])value.Clone();
            byte[] result = new byte[length];
            Buffer.BlockCopy(value, 0, result, 0, length);
            return result;
        }

        /// <summary>取数组中指定偏移与长度的片段。Takes a slice at an offset with a given length.</summary>
        private static byte[] Slice(byte[] value, int offset, int length)
        {
            if (value == null || offset >= value.Length) return new byte[0];
            int count = Math.Min(length, value.Length - offset);
            byte[] result = new byte[count];
            Buffer.BlockCopy(value, offset, result, 0, count);
            return result;
        }

        /// <summary>拼接多段字节。Concatenates several byte arrays.</summary>
        private static byte[] Concat(params byte[][] parts)
        {
            int total = 0;
            for (int i = 0; i < parts.Length; i++) total += parts[i] == null ? 0 : parts[i].Length;
            byte[] result = new byte[total];
            int offset = 0;
            for (int i = 0; i < parts.Length; i++)
            {
                if (parts[i] == null) continue;
                Buffer.BlockCopy(parts[i], 0, result, offset, parts[i].Length);
                offset += parts[i].Length;
            }
            return result;
        }
    }
}
