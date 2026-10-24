using System;
using System.Security.Cryptography;

namespace Wisdom.Pdf
{
    /// <summary>PDF 加密使用的底层算法封装：RC4 与 AES-CBC。Low-level algorithm helpers used by PDF encryption: RC4 and AES-CBC.</summary>
    internal static class PdfCrypto
    {
        /// <summary>
        /// RC4 对称变换（加密与解密同一操作）。
        /// RC4 symmetric transform (encryption and decryption are the same operation).
        /// </summary>
        /// <param name="key">密钥。The key.</param>
        /// <param name="data">输入字节。The input bytes.</param>
        /// <returns>变换后的字节。The transformed bytes.</returns>
        public static byte[] Rc4(byte[] key, byte[] data)
        {
            if (key == null || key.Length == 0) return (byte[])data.Clone();
            byte[] output = new byte[data.Length];
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
        /// <param name="parts">按顺序拼接的输入。Inputs concatenated in order.</param>
        /// <returns>16 字节摘要。The 16-byte digest.</returns>
        public static byte[] Md5(params byte[][] parts)
        {
            using (MD5 md5 = MD5.Create())
            {
                return Hash(md5, parts);
            }
        }

        /// <summary>计算 SHA-256。Computes SHA-256.</summary>
        /// <param name="parts">按顺序拼接的输入。Inputs concatenated in order.</param>
        /// <returns>摘要。The digest.</returns>
        public static byte[] Sha256(params byte[][] parts)
        {
            using (SHA256 sha = SHA256.Create())
            {
                return Hash(sha, parts);
            }
        }

        /// <summary>计算 SHA-384。Computes SHA-384.</summary>
        /// <param name="parts">按顺序拼接的输入。Inputs concatenated in order.</param>
        /// <returns>摘要。The digest.</returns>
        public static byte[] Sha384(params byte[][] parts)
        {
            using (SHA384 sha = SHA384.Create())
            {
                return Hash(sha, parts);
            }
        }

        /// <summary>计算 SHA-512。Computes SHA-512.</summary>
        /// <param name="parts">按顺序拼接的输入。Inputs concatenated in order.</param>
        /// <returns>摘要。The digest.</returns>
        public static byte[] Sha512(params byte[][] parts)
        {
            using (SHA512 sha = SHA512.Create())
            {
                return Hash(sha, parts);
            }
        }

        /// <summary>把多个片段送入哈希算法。Feeds several segments into a hash algorithm.</summary>
        /// <param name="hash">哈希算法。The hash algorithm.</param>
        /// <param name="parts">片段。The segments.</param>
        /// <returns>摘要。The digest.</returns>
        private static byte[] Hash(HashAlgorithm hash, byte[][] parts)
        {
            int total = 0;
            if (parts != null)
            {
                for (int i = 0; i < parts.Length; i++)
                {
                    if (parts[i] != null) total += parts[i].Length;
                }
            }
            byte[] buffer = new byte[total];
            int position = 0;
            if (parts != null)
            {
                for (int i = 0; i < parts.Length; i++)
                {
                    if (parts[i] == null) continue;
                    Buffer.BlockCopy(parts[i], 0, buffer, position, parts[i].Length);
                    position += parts[i].Length;
                }
            }
            return hash.ComputeHash(buffer);
        }

        /// <summary>
        /// AES-CBC 加密并前置 16 字节随机 IV，使用 PKCS#7 填充（PDF 加密的标准形式）。
        /// AES-CBC encryption prefixed with a random 16-byte IV, using PKCS#7 padding (the standard PDF form).
        /// </summary>
        /// <param name="key">密钥。The key.</param>
        /// <param name="data">明文。The plaintext.</param>
        /// <returns>IV 加密文。The IV followed by the ciphertext.</returns>
        public static byte[] AesEncrypt(byte[] key, byte[] data)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;
                aes.GenerateIV();
                byte[] iv = aes.IV;
                using (ICryptoTransform encryptor = aes.CreateEncryptor())
                {
                    byte[] encrypted = encryptor.TransformFinalBlock(data, 0, data.Length);
                    byte[] result = new byte[iv.Length + encrypted.Length];
                    Buffer.BlockCopy(iv, 0, result, 0, iv.Length);
                    Buffer.BlockCopy(encrypted, 0, result, iv.Length, encrypted.Length);
                    return result;
                }
            }
        }

        /// <summary>
        /// AES-CBC 解密（输入前 16 字节为 IV，末尾按 PKCS#7 填充）。
        /// AES-CBC decryption (the first 16 bytes are the IV, the tail is PKCS#7 padded).
        /// </summary>
        /// <param name="key">密钥。The key.</param>
        /// <param name="data">IV 加密文。The IV followed by the ciphertext.</param>
        /// <returns>明文。The plaintext.</returns>
        /// <remarks>填充按宽容方式剥离：规范允许最后一个填充块写得随意，写出器也确实会漏掉补齐或填进非法字节；
        /// 若交给框架校验，一个畸形对象就会让整份文档读不出来。剥离失败时保留原文，交给上层解析器容错。
        /// Padding is stripped leniently: the specification leaves the final padding block free, and writers do
        /// emit payloads that are unaligned or badly padded — strict validation would turn one malformed object
        /// into an unreadable document. When the padding does not parse, the plaintext is kept as is.</remarks>
        public static byte[] AesDecrypt(byte[] key, byte[] data)
        {
            if (data == null || data.Length < 16) return new byte[0];
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.None;
                byte[] iv = new byte[16];
                Buffer.BlockCopy(data, 0, iv, 0, 16);
                aes.IV = iv;
                int start = 16;
                int length = data.Length - 16;
                // 末尾不足 16 字节的残块无法解密，只能丢弃：CBC 的完整块仍然给出全部可用明文。
                // A trailing partial block cannot be decrypted, so it is dropped: the whole blocks still
                // yield every usable plaintext byte.
                length -= length % 16;
                if (length <= 0) return new byte[0];
                byte[] cipher = new byte[length];
                Buffer.BlockCopy(data, start, cipher, 0, length);
                byte[] plain;
                using (ICryptoTransform decryptor = aes.CreateDecryptor())
                {
                    plain = decryptor.TransformFinalBlock(cipher, 0, cipher.Length);
                }
                return StripPkcs7Padding(plain);
            }
        }

        /// <summary>剥离 PKCS#7 填充；填充值不合法或越界时原样返回。Strips PKCS#7 padding; returns the input unchanged when the padding is invalid or overlong.</summary>
        /// <param name="plain">解密后的明文。The decrypted bytes.</param>
        /// <returns>去掉填充的明文。The plaintext without padding.</returns>
        private static byte[] StripPkcs7Padding(byte[] plain)
        {
            if (plain.Length == 0) return plain;
            int count = plain[plain.Length - 1];
            // 先排除 0 与超过整块的填充值，再用固定长度比较：越界的字节数会让逐字节校验跑向负下标。
            // Rule out 0 and a count larger than one block first; a byte-by-byte check alone could walk
            // below the array start.
            if (count <= 0 || count > 16 || count > plain.Length) return plain;
            for (int i = plain.Length - count; i < plain.Length; i++)
            {
                if (plain[i] != (byte)count) return plain;
            }
            byte[] result = new byte[plain.Length - count];
            Buffer.BlockCopy(plain, 0, result, 0, result.Length);
            return result;
        }

        /// <summary>无填充、无 IV 的 AES-CBC 加密，用于 R6 的 UE/OE 与 Perms。Padding-free AES-CBC without an IV, used for R6 UE/OE and Perms.</summary>
        /// <param name="key">密钥。The key.</param>
        /// <param name="iv">初始向量。The initialization vector.</param>
        /// <param name="data">输入（长度须为 16 的倍数）。The input, whose length must be a multiple of 16.</param>
        /// <returns>密文。The ciphertext.</returns>
        public static byte[] AesEncryptNoPadding(byte[] key, byte[] iv, byte[] data)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.None;
                aes.IV = iv;
                using (ICryptoTransform encryptor = aes.CreateEncryptor())
                {
                    return encryptor.TransformFinalBlock(data, 0, data.Length);
                }
            }
        }

        /// <summary>无填充 AES-CBC 解密。Padding-free AES-CBC decryption.</summary>
        /// <param name="key">密钥。The key.</param>
        /// <param name="iv">初始向量。The initialization vector.</param>
        /// <param name="data">密文。The ciphertext.</param>
        /// <returns>明文。The plaintext.</returns>
        public static byte[] AesDecryptNoPadding(byte[] key, byte[] iv, byte[] data)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.None;
                aes.IV = iv;
                using (ICryptoTransform decryptor = aes.CreateDecryptor())
                {
                    return decryptor.TransformFinalBlock(data, 0, data.Length);
                }
            }
        }

        /// <summary>无填充 AES-ECB 加密，用于 /Perms。Padding-free AES-ECB encryption, used by /Perms.</summary>
        /// <param name="key">密钥。The key.</param>
        /// <param name="data">输入。The input.</param>
        /// <returns>密文。The ciphertext.</returns>
        public static byte[] AesEcbEncrypt(byte[] key, byte[] data)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.Mode = CipherMode.ECB;
                aes.Padding = PaddingMode.None;
                using (ICryptoTransform encryptor = aes.CreateEncryptor())
                {
                    return encryptor.TransformFinalBlock(data, 0, data.Length);
                }
            }
        }

        /// <summary>无填充 AES-ECB 解密。Padding-free AES-ECB decryption.</summary>
        /// <param name="key">密钥。The key.</param>
        /// <param name="data">密文。The ciphertext.</param>
        /// <returns>明文。The plaintext.</returns>
        public static byte[] AesEcbDecrypt(byte[] key, byte[] data)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.Mode = CipherMode.ECB;
                aes.Padding = PaddingMode.None;
                using (ICryptoTransform decryptor = aes.CreateDecryptor())
                {
                    return decryptor.TransformFinalBlock(data, 0, data.Length);
                }
            }
        }

        /// <summary>生成密码学随机字节。Generates cryptographically random bytes.</summary>
        /// <param name="length">长度。The length.</param>
        /// <returns>随机字节。The random bytes.</returns>
        public static byte[] Random(int length)
        {
            byte[] bytes = new byte[length];
            using (RandomNumberGenerator random = RandomNumberGenerator.Create())
            {
                random.GetBytes(bytes);
            }
            return bytes;
        }

        /// <summary>把 32 位整数写成小端字节。Writes a 32-bit integer in little-endian order.</summary>
        /// <param name="value">数值。The value.</param>
        /// <returns>4 字节。Four bytes.</returns>
        public static byte[] LittleEndian(int value)
        {
            return new byte[]
            {
                (byte)(value & 0xFF),
                (byte)((value >> 8) & 0xFF),
                (byte)((value >> 16) & 0xFF),
                (byte)((value >> 24) & 0xFF)
            };
        }

        /// <summary>比较两段字节是否相等。Compares two byte arrays for equality.</summary>
        /// <param name="a">第一段。The first array.</param>
        /// <param name="b">第二段。The second array.</param>
        /// <returns>相等返回 true。True when equal.</returns>
        public static bool BytesEqual(byte[] a, byte[] b)
        {
            if (a == null || b == null || a.Length != b.Length) return false;
            int difference = 0;
            for (int i = 0; i < a.Length; i++) difference |= a[i] ^ b[i];
            return difference == 0;
        }
    }
}
