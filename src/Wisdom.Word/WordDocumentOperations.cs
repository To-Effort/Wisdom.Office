using System;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 文档级组合操作：深拷贝、追加与合并，风格参考 Aspose.Words 的 Document.Clone / AppendDocument。
    /// Document-level composition operations: deep clone, append and merge, modelled on Aspose.Words
    /// Document.Clone / AppendDocument.
    /// </summary>
    /// <remarks>
    /// 追加与合并以「节」为单位保留源文档的页面设置、页眉页脚与节分隔，内容会完整深拷贝，
    /// 因此之后修改合并结果不会影响源文档。
    /// Append and merge preserve each source section's page setup, headers, footers and section breaks. All
    /// content is deep-copied, so later edits to the result never affect the source documents.
    /// </remarks>
    public static class WordDocumentOperations
    {
        /// <summary>
        /// 深拷贝整个文档：元数据、默认格式、各节内容、页面设置与页眉页脚全部复制。
        /// Deep-copies a whole document: metadata, default format, section content, page setup and
        /// headers/footers are all duplicated.
        /// </summary>
        /// <param name="document">待拷贝的文档，不能为 null。The document to copy; must not be null.</param>
        /// <returns>独立的新文档。An independent new document.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="document"/> 为 null。Thrown when <paramref name="document"/> is null.</exception>
        public static OfficeDocument Clone(OfficeDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");
            return DocumentCloner.Clone(document);
        }

        /// <summary>
        /// 把源文档的全部节深拷贝追加到目标文档末尾；目标文档的元数据与已有内容保持不变。
        /// Deep-copies every section of the source onto the end of the target; the target metadata and
        /// existing content stay unchanged.
        /// </summary>
        /// <param name="target">目标文档，不能为 null。The target document; must not be null.</param>
        /// <param name="source">源文档，不能为 null。The source document; must not be null.</param>
        /// <returns>追加的节数量。The number of appended sections.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="target"/> 或 <paramref name="source"/> 为 null。Thrown when <paramref name="target"/> or <paramref name="source"/> is null.</exception>
        public static int Append(OfficeDocument target, OfficeDocument source)
        {
            if (target == null) throw new ArgumentNullException("target");
            if (source == null) throw new ArgumentNullException("source");

            for (int i = 0; i < source.Sections.Count; i++)
            {
                Section sourceSection = source.Sections[i];
                PageSetup setup = sourceSection.PageSetup == null ? new PageSetup() : sourceSection.PageSetup.Clone();
                Section section = target.AddSection(setup);
                DocumentCloner.CloneBlocks(sourceSection.Blocks, section.Blocks);
                if (sourceSection.Header != null)
                {
                    DocumentCloner.CloneBlocks(sourceSection.Header.Blocks, section.Header.Blocks);
                }
                if (sourceSection.Footer != null)
                {
                    DocumentCloner.CloneBlocks(sourceSection.Footer.Blocks, section.Footer.Blocks);
                }
            }
            return source.Sections.Count;
        }

        /// <summary>
        /// 按给定顺序把多个文档合并为一个新文档：第一节来自首个非空文档，其后各文档依次深拷贝追加。
        /// 空文档（null）被跳过；全部为空时返回空文档。
        /// Merges several documents into a new document in the given order: the first section set comes from the
        /// first non-null document and the others are deep-copied after it. Null documents are skipped; when
        /// every argument is null the result is an empty document.
        /// </summary>
        /// <param name="documents">待合并的文档列表。The documents to merge.</param>
        /// <returns>合并后的新文档。The merged new document.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="documents"/> 为 null。Thrown when <paramref name="documents"/> is null.</exception>
        public static OfficeDocument Merge(params OfficeDocument[] documents)
        {
            if (documents == null) throw new ArgumentNullException("documents");

            OfficeDocument merged = null;
            for (int i = 0; i < documents.Length; i++)
            {
                if (documents[i] == null) continue;
                if (merged == null)
                {
                    merged = DocumentCloner.Clone(documents[i]);
                    continue;
                }
                Append(merged, documents[i]);
            }
            return merged ?? new OfficeDocument();
        }
    }
}
