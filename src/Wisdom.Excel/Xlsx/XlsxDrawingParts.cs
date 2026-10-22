using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Excel
{
    /// <summary>
    /// 收集工作簿里的图表与图片，生成 OOXML 的 chart / drawing / media 部件及其关系文件。
    /// Collects the charts and images of a workbook and produces the OOXML chart, drawing and
    /// media parts together with their relationship parts.
    /// </summary>
    internal sealed class XlsxDrawingParts
    {
        /// <summary>图表命名空间（drawingml/2006/chart）。 Chart namespace (drawingml/2006/chart).</summary>
        private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        /// <summary>绘图命名空间（drawingml/2006/main）。 Drawing namespace (drawingml/2006/main).</summary>
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        /// <summary>电子表格绘图命名空间（spreadsheetDrawing）。 Spreadsheet drawing namespace (spreadsheetDrawing).</summary>
        private static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        /// <summary>图表部件的关系类型。 Relationship type of a chart part.</summary>
        private const string ChartRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
        /// <summary>图片部件的关系类型。 Relationship type of an image part.</summary>
        private const string ImageRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        /// <summary>工作表到绘图部件的关系类型。 Relationship type from a worksheet to its drawing part.</summary>
        private const string DrawingRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
        /// <summary>图表内容类型。 Content type of a chart part.</summary>
        public const string ContentTypeChart = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
        /// <summary>绘图内容类型。 Content type of a drawing part.</summary>
        public const string ContentTypeDrawing = "application/vnd.openxmlformats-officedocument.drawing+xml";
        /// <summary>一磅对应的 EMU 数值。 EMU units per point.</summary>
        private const double EmuPerPoint = 12700.0;

        /// <summary>待写入 zip 的附加部件（路径 → 字节）。 Extra parts to write into the zip (path to bytes).</summary>
        private readonly List<KeyValuePair<string, byte[]>> _entries = new List<KeyValuePair<string, byte[]>>();
        /// <summary>各表绘图部件的相对目标路径（键为表下标）。 Relative drawing target of each sheet, keyed by sheet index.</summary>
        private readonly Dictionary<int, string> _drawingTargets = new Dictionary<int, string>();
        /// <summary>各表绘图部件的关系 XML（键为表下标）。 Relationship XML of each sheet's drawing part, keyed by sheet index.</summary>
        private readonly Dictionary<int, byte[]> _sheetRelationships = new Dictionary<int, byte[]>();
        /// <summary>已登记的图片扩展名集合。 Registered image file extensions.</summary>
        private readonly HashSet<string> _mediaExtensions = new HashSet<string>(StringComparer.Ordinal);
        /// <summary>图表部件计数（决定 chartN.xml 编号）。 Chart part counter driving chartN.xml numbering.</summary>
        private int _chartIndex;
        /// <summary>图片部件计数（决定 imageN 编号）。 Image counter driving imageN numbering.</summary>
        private int _imageIndex;
        /// <summary>绘图部件计数（决定 drawingN.xml 编号）。 Drawing part counter driving drawingN.xml numbering.</summary>
        private int _drawingIndex;

        /// <summary>附加部件列表。 The extra part entries.</summary>
        public List<KeyValuePair<string, byte[]>> Entries { get { return _entries; } }
        /// <summary>图片扩展名集合，用于写入内容类型默认项。 Image extensions used for the content-type defaults.</summary>
        public HashSet<string> MediaExtensions { get { return _mediaExtensions; } }

        /// <summary>取得指定工作表的绘图部件相对路径；无图表与图片时返回 null。 Gets the relative drawing path of a worksheet, or null when it has neither charts nor images.</summary>
        /// <param name="sheetIndex">表下标（从 0 开始）。 Sheet index (starting from 0).</param>
        /// <returns>相对路径或 null。 The relative target or null.</returns>
        public string GetDrawingTarget(int sheetIndex)
        {
            string target;
            return _drawingTargets.TryGetValue(sheetIndex, out target) ? target : null;
        }

        /// <summary>取得指定工作表需要写出的 .rels 内容；无绘图部件时返回 null。 Gets the .rels content a worksheet needs, or null when it has no drawing part.</summary>
        /// <param name="sheetIndex">表下标（从 0 开始）。 Sheet index (starting from 0).</param>
        /// <returns>关系部件字节或 null。 The relationship part bytes or null.</returns>
        public byte[] GetSheetRelationships(int sheetIndex)
        {
            byte[] data;
            return _sheetRelationships.TryGetValue(sheetIndex, out data) ? data : null;
        }

        /// <summary>扫描全部工作表，生成图表、图片、绘图与关系部件。 Scans every worksheet and builds the chart, image, drawing and relationship parts.</summary>
        /// <param name="workbook">待写出的工作簿。 The workbook being written.</param>
        /// <returns>收集结果。 The collected parts.</returns>
        public static XlsxDrawingParts Build(OfficeWorkbook workbook)
        {
            XlsxDrawingParts parts = new XlsxDrawingParts();
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                if (sheet.Charts.Count == 0 && sheet.Images.Count == 0) continue;
                parts.BuildSheet(workbook, sheet, i);
            }
            return parts;
        }

        /// <summary>为单个工作表生成绘图部件、其关系部件，以及该表的 sheet .rels。 Builds one sheet's drawing part, its relationships and the sheet .rels part.</summary>
        /// <param name="workbook">工作簿。 Workbook.</param>
        /// <param name="sheet">工作表。 Worksheet.</param>
        /// <param name="sheetIndex">表下标。 Sheet index.</param>
        private void BuildSheet(OfficeWorkbook workbook, Worksheet sheet, int sheetIndex)
        {
            _drawingIndex++;
            string drawingName = "xl/drawings/drawing" + _drawingIndex.ToString(CultureInfo.InvariantCulture) + ".xml";
            XElement drawingRoot = new XElement(Xdr + "wsDr",
                new XAttribute(XNamespace.Xmlns + "xdr", Xdr.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName));
            XElement relsRoot = new XElement(Xlsx.PackageRels + "Relationships");
            int relationshipId = 1;
            int shapeId = 2;

            for (int i = 0; i < sheet.Charts.Count; i++)
            {
                ExcelChart chart = sheet.Charts[i];
                _chartIndex++;
                string chartName = "../charts/chart" + _chartIndex.ToString(CultureInfo.InvariantCulture) + ".xml";
                _entries.Add(new KeyValuePair<string, byte[]>(
                    "xl/charts/chart" + _chartIndex.ToString(CultureInfo.InvariantCulture) + ".xml",
                    Xlsx.SerializeXml(XlsxChartParts.Build(chart, workbook, sheet))));
                string id = "rId" + relationshipId.ToString(CultureInfo.InvariantCulture);
                relationshipId++;
                relsRoot.Add(Relationship(id, ChartRelationship, chartName));
                drawingRoot.Add(CreateChartAnchor(chart, id, shapeId));
                shapeId++;
            }

            for (int i = 0; i < sheet.Images.Count; i++)
            {
                ExcelImage image = sheet.Images[i];
                _imageIndex++;
                string extension = string.IsNullOrEmpty(image.Image.FileExtension) ? ".png" : image.Image.FileExtension;
                string mediaPath = "xl/media/image" + _imageIndex.ToString(CultureInfo.InvariantCulture) + extension;
                _entries.Add(new KeyValuePair<string, byte[]>(mediaPath, image.Image.Data));
                _mediaExtensions.Add(extension.TrimStart('.'));
                string id = "rId" + relationshipId.ToString(CultureInfo.InvariantCulture);
                relationshipId++;
                relsRoot.Add(Relationship(id, ImageRelationship, "../media/image"
                    + _imageIndex.ToString(CultureInfo.InvariantCulture) + extension));
                drawingRoot.Add(CreateImageAnchor(image, id, shapeId));
                shapeId++;
            }

            _entries.Add(new KeyValuePair<string, byte[]>(drawingName,
                Xlsx.SerializeXml(new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), drawingRoot))));
            _entries.Add(new KeyValuePair<string, byte[]>(
                "xl/drawings/_rels/drawing" + _drawingIndex.ToString(CultureInfo.InvariantCulture) + ".xml.rels",
                Xlsx.SerializeXml(new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), relsRoot))));

            _drawingTargets[sheetIndex] = "../drawings/drawing" + _drawingIndex.ToString(CultureInfo.InvariantCulture) + ".xml";
            XElement sheetRels = new XElement(Xlsx.PackageRels + "Relationships",
                Relationship("rId1", DrawingRelationship, _drawingTargets[sheetIndex]));
            _sheetRelationships[sheetIndex] = Xlsx.SerializeXml(
                new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), sheetRels));
        }

        /// <summary>把图表、绘图与图片部件的类型写入 [Content_Types].xml。 Writes the chart, drawing and media content types into [Content_Types].xml.</summary>
        /// <param name="root">Types 根元素。 The Types root element.</param>
        public void AddContentTypes(XElement root)
        {
            for (int i = 0; i < _entries.Count; i++)
            {
                string path = _entries[i].Key;
                if (path.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
                if (path.StartsWith("xl/charts/", StringComparison.OrdinalIgnoreCase))
                {
                    root.Add(new XElement(Xlsx.ContentTypes + "Override",
                        new XAttribute("PartName", "/" + path),
                        new XAttribute("ContentType", ContentTypeChart)));
                }
                else if (path.StartsWith("xl/drawings/", StringComparison.OrdinalIgnoreCase))
                {
                    root.Add(new XElement(Xlsx.ContentTypes + "Override",
                        new XAttribute("PartName", "/" + path),
                        new XAttribute("ContentType", ContentTypeDrawing)));
                }
            }
            foreach (string extension in _mediaExtensions)
            {
                root.Add(new XElement(Xlsx.ContentTypes + "Default",
                    new XAttribute("Extension", extension),
                    new XAttribute("ContentType", ImageContentType(extension))));
            }
        }

        /// <summary>按扩展名给出图片 MIME 类型。 Maps an extension to its image MIME type.</summary>
        /// <param name="extension">不含点的扩展名。 Extension without the dot.</param>
        /// <returns>MIME 类型。 The MIME type.</returns>
        private static string ImageContentType(string extension)
        {
            if (string.Equals(extension, "jpg", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, "jpeg", StringComparison.OrdinalIgnoreCase)) return "image/jpeg";
            if (string.Equals(extension, "gif", StringComparison.OrdinalIgnoreCase)) return "image/gif";
            if (string.Equals(extension, "bmp", StringComparison.OrdinalIgnoreCase)) return "image/bmp";
            return "image/png";
        }

        /// <summary>构造一条 OPC 关系元素。 Builds one OPC relationship element.</summary>
        /// <param name="id">关系 Id。 Relationship Id.</param>
        /// <param name="type">关系类型 URL。 Relationship type URL.</param>
        /// <param name="target">目标部件路径。 Target part path.</param>
        /// <returns>Relationship 元素。 The Relationship element.</returns>
        private static XElement Relationship(string id, string type, string target)
        {
            return new XElement(Xlsx.PackageRels + "Relationship",
                new XAttribute("Id", id),
                new XAttribute("Type", type),
                new XAttribute("Target", target));
        }

        /// <summary>生成一个图表在绘图部件中的 oneCellAnchor 图框。 Builds the oneCellAnchor graphic frame of one chart inside the drawing part.</summary>
        /// <param name="chart">图表。 Chart.</param>
        /// <param name="relationshipId">指向 chart 部件的关系 Id。 Relationship Id pointing at the chart part.</param>
        /// <param name="shapeId">形状 Id（绘图内唯一）。 Shape Id, unique inside the drawing.</param>
        /// <returns>oneCellAnchor 元素。 The oneCellAnchor element.</returns>
        private static XElement CreateChartAnchor(ExcelChart chart, string relationshipId, int shapeId)
        {
            return new XElement(Xdr + "oneCellAnchor",
                CreateFrom(chart.Row, chart.Column, 0, 0),
                new XElement(Xdr + "ext",
                    new XAttribute("cx", ToEmu(chart.WidthPoints).ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("cy", ToEmu(chart.HeightPoints).ToString(CultureInfo.InvariantCulture))),
                new XElement(Xdr + "graphicFrame",
                    new XAttribute("macro", string.Empty),
                    new XElement(Xdr + "nvGraphicFramePr",
                        new XElement(Xdr + "cNvPr",
                            new XAttribute("id", shapeId.ToString(CultureInfo.InvariantCulture)),
                            new XAttribute("name", "Chart " + shapeId.ToString(CultureInfo.InvariantCulture))),
                        new XElement(Xdr + "cNvGraphicFramePr")),
                    new XElement(Xdr + "xfrm",
                        new XElement(A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                        new XElement(A + "ext", new XAttribute("cx", "0"), new XAttribute("cy", "0"))),
                    new XElement(A + "graphic",
                        new XElement(A + "graphicData",
                            new XAttribute("uri", C.NamespaceName),
                            new XElement(C + "chart",
                                new XAttribute(XNamespace.Xmlns + "c", C.NamespaceName),
                                new XAttribute(XNamespace.Xmlns + "r", Xlsx.R.NamespaceName),
                                new XAttribute(Xlsx.R + "id", relationshipId))))),
                new XElement(Xdr + "clientData"));
        }

        /// <summary>生成一张图片在绘图部件中的 oneCellAnchor 图片元素。 Builds the oneCellAnchor picture element of one image inside the drawing part.</summary>
        /// <param name="image">锚定图片。 Anchored image.</param>
        /// <param name="relationshipId">指向 media 部件的关系 Id。 Relationship Id pointing at the media part.</param>
        /// <param name="shapeId">形状 Id。 Shape Id.</param>
        /// <returns>oneCellAnchor 元素。 The oneCellAnchor element.</returns>
        private static XElement CreateImageAnchor(ExcelImage image, string relationshipId, int shapeId)
        {
            return new XElement(Xdr + "oneCellAnchor",
                CreateFrom(image.Row, image.Column, image.OffsetXPoints, image.OffsetYPoints),
                new XElement(Xdr + "ext",
                    new XAttribute("cx", ToEmu(image.WidthPoints).ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("cy", ToEmu(image.HeightPoints).ToString(CultureInfo.InvariantCulture))),
                new XElement(Xdr + "pic",
                    new XElement(Xdr + "nvPicPr",
                        new XElement(Xdr + "cNvPr",
                            new XAttribute("id", shapeId.ToString(CultureInfo.InvariantCulture)),
                            new XAttribute("name", "Picture " + shapeId.ToString(CultureInfo.InvariantCulture))),
                        new XElement(Xdr + "cNvPicPr",
                            new XElement(A + "picLocks", new XAttribute("noChangeAspect", "1")))),
                    new XElement(Xdr + "blipFill",
                        new XElement(A + "blip",
                            new XAttribute(XNamespace.Xmlns + "r", Xlsx.R.NamespaceName),
                            new XAttribute(Xlsx.R + "embed", relationshipId)),
                        new XElement(A + "stretch", new XElement(A + "fillRect"))),
                    new XElement(Xdr + "spPr",
                        new XElement(A + "xfrm",
                            new XElement(A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                            new XElement(A + "ext",
                                new XAttribute("cx", ToEmu(image.WidthPoints).ToString(CultureInfo.InvariantCulture)),
                                new XAttribute("cy", ToEmu(image.HeightPoints).ToString(CultureInfo.InvariantCulture)))),
                        new XElement(A + "prstGeom",
                            new XAttribute("prst", "rect"), new XElement(A + "avLst")))),
                new XElement(Xdr + "clientData"));
        }

        /// <summary>生成锚点 from 元素（1 基行列换算为 0 基索引加 EMU 偏移）。 Builds the anchor from element, converting 1-based row/column into zero-based indices plus EMU offsets.</summary>
        /// <param name="row">1 基行号。 One-based row.</param>
        /// <param name="column">1 基列号。 One-based column.</param>
        /// <param name="offsetXPoints">水平偏移（磅）。 Horizontal offset in points.</param>
        /// <param name="offsetYPoints">垂直偏移（磅）。 Vertical offset in points.</param>
        /// <returns>from 元素。 The from element.</returns>
        private static XElement CreateFrom(int row, int column, double offsetXPoints, double offsetYPoints)
        {
            int safeRow = row < 1 ? 1 : row;
            int safeColumn = column < 1 ? 1 : column;
            return new XElement(Xdr + "from",
                new XElement(Xdr + "col", (safeColumn - 1).ToString(CultureInfo.InvariantCulture)),
                new XElement(Xdr + "colOff", ToEmu(offsetXPoints).ToString(CultureInfo.InvariantCulture)),
                new XElement(Xdr + "row", (safeRow - 1).ToString(CultureInfo.InvariantCulture)),
                new XElement(Xdr + "rowOff", ToEmu(offsetYPoints).ToString(CultureInfo.InvariantCulture)));
        }

        /// <summary>把磅值换算为 EMU。 Converts points into EMU.</summary>
        /// <param name="points">磅值。 Point value.</param>
        /// <returns>EMU 整数值。 The EMU integer value.</returns>
        private static long ToEmu(double points)
        {
            if (points <= 0 || double.IsNaN(points) || double.IsInfinity(points)) return 0;
            return (long)Math.Round(points * EmuPerPoint);
        }
    }
}
