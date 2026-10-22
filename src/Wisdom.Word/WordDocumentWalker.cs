using System;
using System.Collections.Generic;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 统一文档模型的只读遍历器：按文档顺序枚举正文、页眉页脚与表格单元格中的段落。
    /// Read-only walker over the unified document model: enumerates paragraphs in body, headers, footers
    /// and table cells in document order.
    /// </summary>
    internal static class WordDocumentWalker
    {
        /// <summary>表格嵌套的最大遍历深度，避免畸形模型导致无限递归。The maximum table nesting depth, preventing infinite recursion on malformed models.</summary>
        private const int MaxTableDepth = 64;

        /// <summary>
        /// 遍历文档中的全部段落。默认先正文后页眉页脚；表格按行、单元格顺序递归。
        /// Walks every paragraph of the document, visiting the body before headers and footers; tables are
        /// visited row by row and cell by cell.
        /// </summary>
        /// <param name="document">待遍历的文档。The document to walk.</param>
        /// <param name="includeHeadersAndFooters">是否包含页眉页脚。Whether headers and footers are included.</param>
        /// <param name="includeTables">是否包含表格单元格内的段落。Whether paragraphs inside table cells are included.</param>
        /// <param name="action">对每个段落的回调。The callback invoked for every paragraph.</param>
        public static void ForEachParagraph(
            OfficeDocument document,
            bool includeHeadersAndFooters,
            bool includeTables,
            Action<Paragraph> action)
        {
            if (document == null || action == null) return;
            for (int s = 0; s < document.Sections.Count; s++)
            {
                Section section = document.Sections[s];
                WalkBlocks(section.Blocks, 0, includeTables, action);
                if (!includeHeadersAndFooters) continue;
                if (section.Header != null) WalkBlocks(section.Header.Blocks, 0, includeTables, action);
                if (section.Footer != null) WalkBlocks(section.Footer.Blocks, 0, includeTables, action);
            }
        }

        /// <summary>
        /// 递归遍历块列表中的段落与表格。
        /// Recursively walks paragraphs and tables in a block list.
        /// </summary>
        /// <param name="blocks">块列表。The block list.</param>
        /// <param name="depth">当前表格嵌套深度。The current table nesting depth.</param>
        /// <param name="includeTables">是否进入表格。Whether tables are entered.</param>
        /// <param name="action">段落回调。The paragraph callback.</param>
        private static void WalkBlocks(List<Block> blocks, int depth, bool includeTables, Action<Paragraph> action)
        {
            if (blocks == null) return;
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    action(paragraph);
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table == null || !includeTables || depth >= MaxTableDepth) continue;
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    TableRow row = table.Rows[r];
                    for (int c = 0; c < row.Cells.Count; c++)
                    {
                        WalkBlocks(row.Cells[c].Blocks, depth + 1, includeTables, action);
                    }
                }
            }
        }
    }
}
