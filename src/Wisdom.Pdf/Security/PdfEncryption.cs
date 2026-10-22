using System;

namespace Wisdom.Pdf
{
    /// <summary>PDF 文档权限位，对应标准安全处理器的 /P 字段。PDF permission flags matching the /P entry of the standard security handler.</summary>
    [Flags]
    public enum PdfPermissions
    {
        /// <summary>不允许任何操作。No operation is allowed.</summary>
        None = 0,
        /// <summary>允许打印。Allows printing.</summary>
        Print = 4,
        /// <summary>允许修改内容。Allows modifying content.</summary>
        Modify = 8,
        /// <summary>允许复制文本与图形。Allows copying text and graphics.</summary>
        Copy = 16,
        /// <summary>允许添加或修改注释。Allows adding or modifying annotations.</summary>
        Annotate = 32,
        /// <summary>允许填写表单域。Allows filling form fields.</summary>
        FillForms = 256,
        /// <summary>允许为无障碍访问提取内容。Allows content extraction for accessibility.</summary>
        Accessibility = 512,
        /// <summary>允许组装页面（插入、旋转、删除）。Allows assembling pages (insert, rotate, delete).</summary>
        Assemble = 1024,
        /// <summary>允许高质量打印。Allows high-quality printing.</summary>
        PrintHighQuality = 2048,
        /// <summary>全部权限。All permissions.</summary>
        All = Print | Modify | Copy | Annotate | FillForms | Accessibility | Assemble | PrintHighQuality
    }

    /// <summary>PDF 加密算法。PDF encryption algorithms.</summary>
    public enum PdfEncryptionAlgorithm
    {
        /// <summary>40 位 RC4（R2）。40-bit RC4 (R2).</summary>
        Rc4_40 = 0,
        /// <summary>128 位 RC4（R3）。128-bit RC4 (R3).</summary>
        Rc4_128 = 1,
        /// <summary>128 位 AES（R4，AESV2）。128-bit AES (R4, AESV2).</summary>
        Aes128 = 2,
        /// <summary>256 位 AES（R6，AESV3），推荐。256-bit AES (R6, AESV3), recommended.</summary>
        Aes256 = 3
    }

    /// <summary>
    /// 加密写出选项。
    /// Options for writing an encrypted document.
    /// </summary>
    public sealed class PdfEncryptionOptions
    {
        /// <summary>创建默认选项：AES-256、全部权限、加密元数据。Creates defaults: AES-256, all permissions, encrypted metadata.</summary>
        public PdfEncryptionOptions()
        {
            Algorithm = PdfEncryptionAlgorithm.Aes256;
            Permissions = PdfPermissions.All;
            EncryptMetadata = true;
        }

        /// <summary>用户口令（打开文档）。The user password used to open the document.</summary>
        public string UserPassword { get; set; }
        /// <summary>所有者口令（修改权限）；为空时使用用户口令。The owner password used to change permissions; the user password when empty.</summary>
        public string OwnerPassword { get; set; }
        /// <summary>加密算法。The encryption algorithm.</summary>
        public PdfEncryptionAlgorithm Algorithm { get; set; }
        /// <summary>权限位。The permission flags.</summary>
        public PdfPermissions Permissions { get; set; }
        /// <summary>是否加密元数据流。Whether metadata streams are encrypted.</summary>
        public bool EncryptMetadata { get; set; }
    }
}
