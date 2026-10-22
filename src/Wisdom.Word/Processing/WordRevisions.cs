using System;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 修订统计与接受/拒绝工具：批量把文档中的插入、删除修订固化为最终内容或还原为原始内容。
    /// Revision statistics and accept/reject helpers that either commit tracked insertions and deletions or
    /// restore the original content.
    /// </summary>
    /// <remarks>
    /// 段落标记修订只清除标记、不合并相邻段落；被删除的内容保存在 <see cref="Paragraph.DeletedInlines"/>，
    /// 接受修订时丢弃、拒绝修订时恢复到正文。
    ///
    /// Paragraph-mark revisions are cleared without merging adjacent paragraphs; deleted content lives in
    /// <see cref="Paragraph.DeletedInlines"/> and is discarded on accept or restored on reject.
    /// </remarks>
    public static class WordRevisions
    {
        /// <summary>
        /// 统计文档中的修订数量：每个带修订标记的行内元素、每段被删除的内容以及每个带标记的段落各计一处。
        /// Counts the revisions in a document: every inline carrying a revision marker, every deleted-content
        /// paragraph fragment and every tracked paragraph mark counts as one change.
        /// </summary>
        /// <param name="document">待统计的文档。The document to inspect.</param>
        /// <returns>修订数量；文档为 null 时返回 0。The revision count; 0 when the document is null.</returns>
        public static int Count(OfficeDocument document)
        {
            if (document == null) return 0;
            int count = 0;
            WordDocumentWalker.ForEachParagraph(document, true, true, delegate (Paragraph paragraph)
            {
                count += CountParagraph(paragraph);
            });
            return count;
        }

        /// <summary>
        /// 接受全部修订：保留插入内容、丢弃被删除内容并清除所有修订标记。
        /// Accepts every revision: keeps insertions, discards deleted content and clears all revision markers.
        /// </summary>
        /// <param name="document">待处理的文档。The document to process.</param>
        /// <returns>被固化的修订数量。The number of revisions committed.</returns>
        public static int AcceptAll(OfficeDocument document)
        {
            if (document == null) return 0;
            int handled = 0;
            WordDocumentWalker.ForEachParagraph(document, true, true, delegate (Paragraph paragraph)
            {
                handled += CountParagraph(paragraph);
                paragraph.DeletedInlines.Clear();
                for (int i = paragraph.Inlines.Count - 1; i >= 0; i--)
                {
                    Inline inline = paragraph.Inlines[i];
                    if (inline.Revision != null && inline.Revision.Type == RevisionType.Delete)
                    {
                        paragraph.Inlines.RemoveAt(i);
                        continue;
                    }
                    inline.Revision = null;
                }
                paragraph.Revision = null;
            });
            return handled;
        }

        /// <summary>
        /// 拒绝全部修订：移除插入内容、把被删除内容恢复到正文并清除所有修订标记。
        /// Rejects every revision: removes insertions, restores deleted content to the body and clears all
        /// revision markers.
        /// </summary>
        /// <remarks>模型不记录被删除内容的原始字符偏移，恢复时统一追加到段落末尾。The model does not track the original character offsets of deleted content, so restored fragments are appended to the end of their paragraph.</remarks>
        /// <param name="document">待处理的文档。The document to process.</param>
        /// <returns>被还原的修订数量。The number of revisions rolled back.</returns>
        public static int RejectAll(OfficeDocument document)
        {
            if (document == null) return 0;
            int handled = 0;
            WordDocumentWalker.ForEachParagraph(document, true, true, delegate (Paragraph paragraph)
            {
                handled += CountParagraph(paragraph);
                for (int i = paragraph.Inlines.Count - 1; i >= 0; i--)
                {
                    Inline inline = paragraph.Inlines[i];
                    if (inline.Revision != null && inline.Revision.Type == RevisionType.Insert)
                    {
                        paragraph.Inlines.RemoveAt(i);
                    }
                }
                for (int i = 0; i < paragraph.DeletedInlines.Count; i++)
                {
                    Inline restored = paragraph.DeletedInlines[i];
                    restored.Revision = null;
                    paragraph.Inlines.Add(restored);
                }
                paragraph.DeletedInlines.Clear();
                for (int i = 0; i < paragraph.Inlines.Count; i++)
                {
                    paragraph.Inlines[i].Revision = null;
                }
                paragraph.Revision = null;
            });
            return handled;
        }

        /// <summary>
        /// 统计单个段落的修订数量（行内标记、删除内容与段落标记）。
        /// Counts the revisions of a single paragraph (inline markers, deleted content and the paragraph mark).
        /// </summary>
        /// <param name="paragraph">待统计的段落。The paragraph to inspect.</param>
        /// <returns>修订数量。The revision count.</returns>
        private static int CountParagraph(Paragraph paragraph)
        {
            if (paragraph == null) return 0;
            int count = paragraph.Revision == null ? 0 : 1;
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                if (paragraph.Inlines[i].Revision != null) count++;
            }
            count += paragraph.DeletedInlines.Count;
            return count;
        }
    }
}
