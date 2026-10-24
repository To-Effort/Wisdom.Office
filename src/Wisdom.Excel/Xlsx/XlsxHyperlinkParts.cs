using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Excel
{
    /// <summary>
    /// 收集工作表中的超链接并生成 hyperlinks 元素与工作表关系部件。外部链接写入 External 关系，
    /// 以 "#" 开头的表内位置写成 location；不安全的外部协议（javascript/data/vbscript 等）会被跳过。
    /// Collects the hyperlinks of every worksheet and produces the hyperlinks element and the sheet
    /// relationship part. External links become External relationships, targets starting with "#" are
    /// written as in-workbook locations, and unsafe schemes (javascript, data, vbscript and the like) are skipped.
    /// </summary>
    internal sealed class XlsxHyperlinkParts
    {
        /// <summary>工作表到超链接的关系类型。Relationship type from a worksheet to a hyperlink.</summary>
        private const string HyperlinkRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        /// <summary>工作表到绘图部件的关系类型。Relationship type from a worksheet to its drawing part.</summary>
        private const string DrawingRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
        /// <summary>Excel 接受的最长超链接地址。The longest hyperlink target Excel accepts.</summary>
        private const int MaxTargetLength = 2079;

        /// <summary>一个待写出的超链接：所属单元格、目标、是否外部链接与关系 Id。One hyperlink to write: its cell, target, external flag and relationship id.</summary>
        internal sealed class Link
        {
            /// <summary>超链接所在的单元格。The cell that owns the hyperlink.</summary>
            public Cell Cell;
            /// <summary>外部 URL 或 "#" 开头的表内位置。The external URL, or the in-workbook location starting with "#".</summary>
            public string Target;
            /// <summary>是否为外部链接。Whether the target is external.</summary>
            public bool External;
            /// <summary>外部链接的关系 Id；表内链接为 null。The relationship id of an external link; null for internal links.</summary>
            public string RelationshipId;
        }

        /// <summary>各表的超链接列表（键为表下标）。Hyperlink lists keyed by sheet index.</summary>
        private readonly Dictionary<int, List<Link>> _links = new Dictionary<int, List<Link>>();

        /// <summary>扫描工作簿，为每个外部链接分配关系 Id（有绘图部件时从 rId2 开始）。Scans the workbook and assigns a relationship id to every external link, starting at rId2 when a drawing part occupies rId1.</summary>
        /// <param name="workbook">待写出的工作簿。The workbook being written.</param>
        /// <param name="decorations">图表与图片部件；可为 null。The chart and image parts; may be null.</param>
        /// <returns>收集结果。The collected hyperlinks.</returns>
        public static XlsxHyperlinkParts Build(OfficeWorkbook workbook, XlsxDrawingParts decorations)
        {
            XlsxHyperlinkParts parts = new XlsxHyperlinkParts();
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                List<Link> links = new List<Link>();
                int relationshipId = decorations != null && decorations.GetDrawingTarget(i) != null ? 2 : 1;
                foreach (Cell cell in workbook.Worksheets[i].Cells)
                {
                    string target = cell.Hyperlink;
                    if (string.IsNullOrEmpty(target)) continue;
                    target = target.Trim();
                    if (target.Length < 2) continue;
                    bool external = target[0] != '#';
                    if (external && !IsSafeTarget(target)) continue;
                    Link link = new Link();
                    link.Cell = cell;
                    link.Target = target;
                    link.External = external;
                    if (external)
                    {
                        link.RelationshipId = "rId" + relationshipId.ToString(CultureInfo.InvariantCulture);
                        relationshipId++;
                    }
                    links.Add(link);
                }
                if (links.Count > 0) parts._links[i] = links;
            }
            return parts;
        }

        /// <summary>
        /// 判断外部链接是否安全且长度合法：拒绝控制字符、超长地址、脚本类协议与指向本机或共享盘的协议。
        /// Checks that an external target is safe and short enough: control characters, overlong targets, script
        /// schemes and schemes that reach a local disk or a network share are rejected.
        /// </summary>
        /// <param name="target">外部链接目标。The external target.</param>
        /// <returns>可写出时返回 true。True when it can be written.</returns>
        /// <remarks>协议白名单与 <c>WordFormatUtil.IsSafeHyperlinkTarget</c> 保持一致：file: 与 smb: 会让收到文件的
        /// 人点开自己并无意识的本地路径或 UNC 共享（Windows 打开 UNC 路径还会顺带发出凭据协商），
        /// 因此不再放行，Markdown 侧早已禁止。
        /// The allow-list mirrors <c>WordFormatUtil.IsSafeHyperlinkTarget</c>: file: and smb: are no longer accepted,
        /// because a recipient can click a path on their own machine or a UNC share they never intended to touch
        /// (opening a UNC path on Windows also triggers credential negotiation), and the Markdown side already
        /// refused them.</remarks>
        private static bool IsSafeTarget(string target)
        {
            if (target.Length > MaxTargetLength) return false;
            for (int i = 0; i < target.Length; i++)
            {
                if (target[i] < 32) return false;
            }
            int colon = target.IndexOf(':');
            int slash = target.IndexOfAny(new char[] { '/', '\\', '?', '#' });
            if (colon < 0 || (slash >= 0 && colon > slash)) return true;
            if (colon == 0) return false;
            string scheme = target.Substring(0, colon).ToLowerInvariant();
            return scheme == "http" || scheme == "https" || scheme == "mailto" || scheme == "ftp"
                || scheme == "ftps" || scheme == "tel" || scheme == "sms";
        }

        /// <summary>判断指定工作表是否含可写出的超链接。Reports whether a sheet has any writable hyperlink.</summary>
        /// <param name="sheetIndex">表下标。Sheet index.</param>
        /// <returns>含超链接返回 true。True when the sheet has hyperlinks.</returns>
        public bool HasLinks(int sheetIndex)
        {
            return _links.ContainsKey(sheetIndex);
        }

        /// <summary>写出工作表的 hyperlinks 元素；该表没有超链接时不写任何内容。Writes the hyperlinks element of a sheet, or nothing when it has no hyperlinks.</summary>
        /// <param name="writer">目标写出器。Target writer.</param>
        /// <param name="sheetIndex">表下标。Sheet index.</param>
        public void Write(XmlWriter writer, int sheetIndex)
        {
            List<Link> links;
            if (!_links.TryGetValue(sheetIndex, out links)) return;
            string main = Xlsx.Main.NamespaceName;
            writer.WriteStartElement("hyperlinks", main);
            for (int i = 0; i < links.Count; i++)
            {
                Link link = links[i];
                writer.WriteStartElement("hyperlink", main);
                writer.WriteAttributeString("ref", CellAddress.Format(link.Cell.Row, link.Cell.Column));
                if (link.External)
                {
                    writer.WriteAttributeString("r", "id", Xlsx.R.NamespaceName, link.RelationshipId);
                }
                else
                {
                    writer.WriteAttributeString("location", link.Target.Substring(1));
                }
                if (!string.IsNullOrEmpty(link.Cell.HyperlinkTooltip))
                {
                    writer.WriteAttributeString("tooltip", link.Cell.HyperlinkTooltip);
                }
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        /// <summary>生成工作表的关系部件；既无绘图也无超链接时返回 null。Builds the worksheet relationship part, returning null when the sheet has neither a drawing nor hyperlinks.</summary>
        /// <param name="sheetIndex">表下标。Sheet index.</param>
        /// <param name="drawingTarget">绘图部件相对路径；无绘图时为 null。The drawing part target, or null.</param>
        /// <returns>关系部件根元素或 null。The relationship root element, or null.</returns>
        public XElement CreateRelationships(int sheetIndex, string drawingTarget)
        {
            List<Link> links;
            _links.TryGetValue(sheetIndex, out links);
            bool hasExternal = false;
            if (links != null)
            {
                for (int i = 0; i < links.Count; i++)
                {
                    if (links[i].External) { hasExternal = true; break; }
                }
            }
            if (drawingTarget == null && !hasExternal) return null;

            XElement root = new XElement(Xlsx.PackageRels + "Relationships");
            if (drawingTarget != null)
            {
                root.Add(new XElement(Xlsx.PackageRels + "Relationship",
                    new XAttribute("Id", "rId1"),
                    new XAttribute("Type", DrawingRelationship),
                    new XAttribute("Target", drawingTarget)));
            }
            if (links != null)
            {
                for (int i = 0; i < links.Count; i++)
                {
                    Link link = links[i];
                    if (!link.External) continue;
                    root.Add(new XElement(Xlsx.PackageRels + "Relationship",
                        new XAttribute("Id", link.RelationshipId),
                        new XAttribute("Type", HyperlinkRelationship),
                        new XAttribute("Target", link.Target),
                        new XAttribute("TargetMode", "External")));
                }
            }
            return root;
        }
    }
}
