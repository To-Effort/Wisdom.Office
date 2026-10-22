using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

internal static class Walker
{
    public static void Run(string path)
    {
        using (ZipArchive zip = ZipFile.OpenRead(path))
        {
            ZipArchiveEntry entry = zip.GetEntry("xl/worksheets/sheet1.xml");
            using (Stream s = entry.Open())
            using (MemoryStream buffer = new MemoryStream())
            {
                s.CopyTo(buffer);
                byte[] xml = buffer.ToArray();
                Console.WriteLine("part bytes=" + xml.Length);
                Console.WriteLine(Encoding.UTF8.GetString(xml, 0, Math.Min(xml.Length, 700)));
                Console.WriteLine("--- node walk ---");
                Walk(xml);
            }
        }

    }

    private static void Walk(byte[] xml)
    {
        XmlReaderSettings settings = new XmlReaderSettings();
        settings.DtdProcessing = DtdProcessing.Prohibit;
        settings.XmlResolver = null;
        settings.IgnoreComments = true;
        settings.IgnoreWhitespace = true;
        int rootDepth = -1;
        int topLevels = 0;
        using (MemoryStream memory = new MemoryStream(xml))
        using (XmlReader reader = XmlReader.Create(memory, settings))
        {
            reader.MoveToContent();
            rootDepth = reader.Depth;
            Console.WriteLine("root=" + reader.LocalName + " depth=" + reader.Depth + " empty=" + reader.IsEmptyElement);
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.Depth != rootDepth + 1) continue;
                topLevels++;
                string name = reader.LocalName;
                if (name == "sheetData")
                {
                    Console.WriteLine("  [" + topLevels + "] sheetData depth=" + reader.Depth + " empty=" + reader.IsEmptyElement);
                    int dataDepth = reader.Depth;
                    int rows = 0;
                    while (reader.Read())
                    {
                        if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == dataDepth) break;
                        if (reader.NodeType != XmlNodeType.Element) continue;
                        if (reader.Depth != dataDepth + 1 || reader.LocalName != "row") continue;
                        rows++;
                        if (reader.IsEmptyElement) continue;
                        int rowDepth = reader.Depth;
                        int cells = 0;
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == rowDepth) break;
                            if (reader.NodeType != XmlNodeType.Element) continue;
                            if (reader.Depth != rowDepth + 1 || reader.LocalName != "c") continue;
                            cells++;
                            if (reader.IsEmptyElement) continue;
                            int cellDepth = reader.Depth;
                            while (reader.Read())
                            {
                                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == cellDepth) break;
                                if (reader.NodeType != XmlNodeType.Element) continue;
                                if (reader.LocalName == "v" || reader.LocalName == "f")
                                {
                                    string text = reader.ReadElementContentAsString();
                                    Console.WriteLine("      v/f='" + reader.LocalName + "' depth-after=" + reader.Depth
                                        + " type-after=" + reader.NodeType + " text=" + text);
                                }
                            }
                        }
                        Console.WriteLine("    row " + rows + " cells=" + cells + " depth-after=" + reader.Depth + " type=" + reader.NodeType);
                        if (rows > 3) { Console.WriteLine("    (truncated)"); break; }
                    }
                    Console.WriteLine("  sheetData rows-seen=" + rows);
                    continue;
                }
                Console.WriteLine("  [" + topLevels + "] " + name + " depth=" + reader.Depth + " empty=" + reader.IsEmptyElement);
                if (!reader.IsEmptyElement)
                {
                    int depth = reader.Depth;
                    int inner = 0;
                    while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth))
                    {
                        if (reader.NodeType == XmlNodeType.Element) inner++;
                    }
                    Console.WriteLine("        inner elements=" + inner + " now: depth=" + reader.Depth + " type=" + reader.NodeType + " name=" + reader.LocalName);
                }
                if (topLevels > 12) break;
            }
            Console.WriteLine("top-levels=" + topLevels);
        }
    }
}
