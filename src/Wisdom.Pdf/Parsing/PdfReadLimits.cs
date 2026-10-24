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
        /// <summary>
        /// 对象号允许的最大值。真实文档常年不超过几万，封顶后 /Size = 最大对象号 + 1 与
        /// 对象号自增都不会溢出 int；畸形文件声明的 2147483647 因此不能再把编辑后的文件写坏。
        /// Maximum object number that may be registered. Real documents stay well under a few
        /// tens of thousands, so capping keeps both "/Size = highest number + 1" and the object
        /// number counter inside int range, and a malformed file claiming 2147483647 can no
        /// longer corrupt the edited output.
        /// </summary>
        public static int MaxObjectNumber = 100000000;
        /// <summary>交叉表中允许解析的最大条目数。Maximum number of cross-reference entries allowed to be parsed.</summary>
        public static int MaxXrefEntries = 2000000;
        /// <summary>单个对象流中允许包含的最大对象数。Maximum number of objects allowed inside a single object stream.</summary>
        public static int MaxObjectStreamObjects = 1000000;
        /// <summary>页面树允许遍历的最大深度，用于阻断循环引用导致的无限递归。Maximum depth the page tree may be traversed; blocks unbounded recursion caused by circular references.</summary>
        public static int MaxPageTreeDepth = 64;
        /// <summary>单个字体编码映射（CMap 与 /W 宽度表）允许登记的最大条目数，超出即丢弃后续区间，
        /// 因为畸形 CMap 流可以用大量合法跨度把映射表堆到内存耗尽。
        /// Maximum entries a single glyph map (CMap or /W width table) may hold; wider sections are dropped,
        /// because a malformed CMap stream can stack many individually legal ranges until memory runs out.</summary>
        public static int MaxGlyphMapEntries = 300000;
        /// <summary>语法对象（数组与字典）允许嵌套的最大层数，用于阻断畸形词元把调用栈打满。
        /// Maximum nesting depth of syntactic objects (arrays and dictionaries); blocks a malformed token
        /// stream from exhausting the call stack.</summary>
        public static int MaxObjectDepth = 64;
    }
}
