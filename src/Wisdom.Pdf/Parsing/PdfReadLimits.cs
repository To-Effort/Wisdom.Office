using System;

namespace Wisdom.Pdf
{
    /// <summary>
    /// 解析不可信 PDF 文件时使用的进程级防护上限，用于限制内存膨胀（解压炸弹）
    /// 和病态结构（超大交叉表、循环引用的页面树）。
    /// Process-wide safety limits used when parsing untrusted PDF files. They bound memory growth
    /// from decompression bombs and pathological structures such as oversized cross-reference
    /// tables and page trees with circular references.
    /// </summary>
    public static class PdfReadLimits
    {
        /// <summary>单个流解码后允许的最大字节数，默认 256 MB，超出即判定为解压炸弹。Maximum decoded byte size allowed for a single stream; defaults to 256 MB and anything larger is treated as a decompression bomb.</summary>
        public static long MaxDecodedStreamBytes = 256L * 1024 * 1024;
        /// <summary>交叉表中允许解析的最大条目数。Maximum number of cross-reference entries allowed to be parsed.</summary>
        public static int MaxXrefEntries = 2000000;
        /// <summary>单个对象流中允许包含的最大对象数。Maximum number of objects allowed inside a single object stream.</summary>
        public static int MaxObjectStreamObjects = 1000000;
        /// <summary>页面树允许遍历的最大深度，用于阻断循环引用导致的无限递归。Maximum depth the page tree may be traversed; blocks unbounded recursion caused by circular references.</summary>
        public static int MaxPageTreeDepth = 64;
    }
}
