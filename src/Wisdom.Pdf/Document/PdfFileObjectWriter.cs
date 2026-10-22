using System;

namespace Wisdom.Pdf
{
    /// <summary>
    /// 把 <see cref="PdfObjectWriter"/> 的登记请求转发到对象级文档 <see cref="PdfFile"/> 的适配器。
    /// Adapter that forwards <see cref="PdfObjectWriter"/> registrations into the object-level
    /// <see cref="PdfFile"/> document.
    /// </summary>
    /// <remarks>
    /// 字体、图像等既有写出流程只依赖 Add/AddStream，因此经本适配器即可把这些对象直接写入已打开的文档，
    /// 并在增量保存时随变更一并写出；对象号由文档统一分配。
    /// The existing font and image writers only rely on Add/AddStream, so this adapter lets them write into an
    /// already opened document, where object numbers come from the document and the objects are saved with the
    /// pending changes; encryption is applied by the document writer as usual.
    /// </remarks>
    internal sealed class PdfFileObjectWriter : PdfObjectWriter
    {
        /// <summary>接收对象的文档。The document receiving the objects.</summary>
        private readonly PdfFile _file;

        /// <summary>
        /// 创建绑定到指定文档的写出适配器。
        /// Creates the writer adapter bound to the given document.
        /// </summary>
        /// <param name="file">目标文档。The target document.</param>
        /// <exception cref="ArgumentNullException"><paramref name="file"/> 为 null 时抛出。Thrown when <paramref name="file"/> is null.</exception>
        public PdfFileObjectWriter(PdfFile file)
        {
            if (file == null) throw new ArgumentNullException("file");
            _file = file;
        }

        /// <summary>
        /// 把对象登记为文档的间接对象。
        /// Registers the object as an indirect object of the document.
        /// </summary>
        /// <param name="value">待登记的对象。The object to register.</param>
        /// <returns>文档分配的引用。The reference assigned by the document.</returns>
        public override PdfReference Add(PdfObject value)
        {
            return _file.AddObjectReference(value);
        }
    }
}
