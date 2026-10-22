using System;
using System.Collections.Generic;
using System.IO;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Pdf;

// PDF 演示：
// 1. 用统一模型经 OfficeConverter 生成多元素 PDF
// 2. PdfDocument 回读：页数、页面尺寸、文本提取、表格识别
// 3. PDF -> DOCX 反向转换并校验
// 4. PdfFile 对象级编辑：合并页面、加水印、增量保存
// 5. AcroForm：拼表单、读字段、批量填写、生成外观、扁平化
// 6. 口令加密：四种算法写出、用户/所有者口令回读、错口令拒绝
internal static class Program
{
    private static string OutDir
    {
        get
        {
            string dir = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "out");
            Directory.CreateDirectory(dir);
            return dir;
        }
    }

    private static void Main()
    {
        Console.WriteLine("=== Wisdom.Sample.Pdf：PDF 生成与解析演示 ===");

        // 1. 生成 PDF
        OfficeDocument document = new OfficeDocument();
        document.Metadata.Title = "PDF 生成示例";

        Section section = document.AddSection(PageSetup.A4());
        section.Blocks.Add(new Paragraph("季度销售汇总").AddText("（自动生成）", new TextFormat("Arial", 9)));
        section.Blocks.Add(new Paragraph("下面是一张三行两列的表格，用于验证 PDF 解析时的表格识别。"));

        Table table = new Table();
        table.Columns.Add(new TableColumn(180));
        table.Columns.Add(new TableColumn(120));
        table.AddRow("产品", "金额");
        table.AddRow("无线键盘", "3,594.00");
        table.AddRow("显示器", "9,495.00");
        section.Blocks.Add(table);

        string pdfPath = Path.Combine(OutDir, "sample-pdf.pdf");
        OfficeConverter.ToPdf(document, pdfPath);
        Console.WriteLine("[写入] PDF: " + new FileInfo(pdfPath).Length + " 字节");

        // 2. 回读解析
        PdfDocument pdf = PdfDocument.Load(pdfPath);
        Console.WriteLine("[回读] 页数=" + pdf.PageCount + "，元数据标题=" + pdf.Info.Title);
        foreach (PdfPage page in pdf.Pages)
        {
            Console.WriteLine("  第 " + page.Number + " 页 " + Math.Round(page.Width, 1) + "x" + Math.Round(page.Height, 1)
                + "，文本=" + Truncate(page.ExtractText(), 30)
                + "，表格=" + page.GetTables().Count + " 个");
        }

        // 3. PDF -> DOCX
        string docxPath = Path.Combine(OutDir, "sample-pdf-roundtrip.docx");
        OfficeConverter.Convert(pdfPath, docxPath);
        OfficeDocument back = Wisdom.Word.DocxReader.Load(docxPath);
        Console.WriteLine("[转换] DOCX: " + new FileInfo(docxPath).Length + " 字节，回读段落块=" + CountBlocks(back));

        RunWatermarkDemo(pdfPath);
        RunFormDemo(pdfPath);
        RunEncryptionDemo(pdfPath);

        Console.WriteLine("PDF 演示完成，输出目录: " + Path.GetFullPath(OutDir));
    }

    /// <summary>加密演示：同一份文档按四种算法写出，再用用户口令与所有者口令分别回读。Encryption demo: write the same document with all four algorithms, then re-open each with either password.</summary>
    /// <param name="pdfPath">上一步生成的 PDF 路径。The PDF written by the previous step.</param>
    private static void RunEncryptionDemo(string pdfPath)
    {
        PdfEncryptionAlgorithm[] algorithms =
        {
            PdfEncryptionAlgorithm.Rc4_40,
            PdfEncryptionAlgorithm.Rc4_128,
            PdfEncryptionAlgorithm.Aes128,
            PdfEncryptionAlgorithm.Aes256
        };

        foreach (PdfEncryptionAlgorithm algorithm in algorithms)
        {
            string path = Path.Combine(OutDir, "pdf-enc-" + algorithm + ".pdf");
            PdfFile file = PdfFile.Load(pdfPath);

            PdfEncryptionOptions options = new PdfEncryptionOptions();
            options.Algorithm = algorithm;
            options.UserPassword = "reader-pwd";
            options.OwnerPassword = "owner-pwd";
            options.Permissions = PdfPermissions.Print | PdfPermissions.Copy;
            file.SetEncryption(options);
            int pages = file.PageCount;
            file.Save(path);

            PdfFile asUser = PdfFile.Load(path, "reader-pwd");
            PdfFile asOwner = PdfFile.Load(path, "owner-pwd");
            string text = PdfDocument.Load(path, "reader-pwd").ExtractText();
            bool rejected = false;
            try
            {
                PdfFile.Load(path, "wrong-pwd");
            }
            catch (PdfPasswordException)
            {
                rejected = true;
            }

            Console.WriteLine("[加密] " + algorithm + " 算法=" + asUser.EncryptionAlgorithm
                + " 页数=" + pages + "/" + asUser.PageCount + "/" + asOwner.PageCount
                + " 提取字符=" + text.Length + " 拒绝错口令=" + rejected
                + "用户权限=" + asUser.Permissions + " 所有者权限=" + asOwner.Permissions);
        }
    }

    /// <summary>表单演示：对象级拼出 AcroForm，再读取字段、批量填写、生成外观与扁平化。Form demo: build an AcroForm at object level, then read the fields, fill them, generate appearances and flatten.</summary>
    /// <param name="pdfPath">上一步生成的 PDF 路径。The PDF written by the previous step.</param>
    private static void RunFormDemo(string pdfPath)
    {
        string blankPath = Path.Combine(OutDir, "pdf-form-blank.pdf");
        string filledPath = Path.Combine(OutDir, "pdf-form-filled.pdf");
        string flatPath = Path.Combine(OutDir, "pdf-form-flattened.pdf");

        // 5.1 现拼一份空白表单：两个文本域、一个组合框、一个复选框
        PdfFile blank = PdfFile.Load(pdfPath);
        AddDemoForm(blank);
        blank.Save(blankPath);

        // 5.2 读取字段清单：名称、类型、当前值
        PdfFile filled = PdfFile.Load(blankPath);
        PdfForm form = filled.GetForm();
        string inventory = string.Empty;
        foreach (PdfFormField field in form.Fields)
        {
            if (field.FullyQualifiedName.Length == 0) continue;
            inventory += field.FullyQualifiedName + "(" + field.Kind + ")" + (string.IsNullOrEmpty(field.Value) ? " " : "=" + Truncate(field.Value, 12) + " ");
        }

        // 5.3 批量填写：中文值会自动嵌入子集字体，外观流按控件矩形排版
        Dictionary<string, string> values = new Dictionary<string, string>();
        values["applicant.name"] = "王小明 Wang Xiaoming";
        values["applicant.address"] = "上海市浦东新区世纪大道 100 号\n100 Century Avenue, Pudong";
        values["applicant.city"] = "shenzhen";
        values["applicant.agree"] = "true";
        int filledCount = filled.FillForm(values);
        filled.Save(filledPath);

        // 5.4 扁平化：值烧进页面内容，控件与 AcroForm 一并消失，文档不再是可填表单
        PdfFile flat = PdfFile.Load(filledPath);
        int flattened = flat.FlattenForm();
        flat.Save(flatPath);

        PdfFile afterFlat = PdfFile.Load(flatPath);
        string extracted = PdfDocument.Load(flatPath).Pages[0].ExtractText();
        Console.WriteLine("[表单] 字段清单=" + Truncate(inventory, 96));
        Console.WriteLine("[表单] 填写=" + filledCount + " 个，扁平化=" + flattened + " 个，扁平后仍有表单="
            + (afterFlat.HasForm ? "是" : "否") + "，首页可读=" + (extracted.Contains("Wang Xiaoming") ? "是" : "否"));
    }

    /// <summary>用对象级 API 给首页拼出一份四字段 AcroForm。Builds a four-field AcroForm on page one through the object-level API.</summary>
    /// <param name="file">目标文件。The target file.</param>
    private static void AddDemoForm(PdfFile file)
    {
        PdfReference page = file.GetPageReference(0);
        PdfDictionary pageDictionary = file.GetPageDictionary(0);
        PdfArray fields = new PdfArray();
        PdfArray annots = new PdfArray();

        fields.Add(AddDemoField(file, page, "applicant.name", "Tx", 0, new double[] { 72, 700, 292, 716 }, annots));
        fields.Add(AddDemoField(file, page, "applicant.address", "Tx", 0x800, new double[] { 72, 660, 292, 694 }, annots));

        PdfReference city = AddDemoField(file, page, "applicant.city", "Ch", 0x20000, new double[] { 320, 700, 440, 716 }, annots);
        PdfDictionary cityDictionary = file.GetObject<PdfDictionary>(city.Number);
        cityDictionary.Set("Opt", new PdfArray()
            .Add(new PdfArray().AddName("beijing").Add(PdfString.Unicode("北京 Beijing")))
            .Add(new PdfArray().AddName("shenzhen").Add(PdfString.Unicode("深圳 Shenzhen")))
            .Add(new PdfArray().AddName("chengdu").Add(PdfString.Unicode("成都 Chengdu"))));
        file.MarkModified(city.Number);
        fields.Add(city);

        // 复选框的勾选态名称来自 /AP /N 的键，因此必须显式给出开与关两个状态
        PdfReference agree = AddDemoField(file, page, "applicant.agree", "Btn", 0x40000, new double[] { 72, 636, 84, 648 }, annots);
        PdfDictionary blank = new PdfDictionary();
        blank.SetName("Type", "XObject");
        blank.SetName("Subtype", "Form");
        blank.Set("BBox", new PdfArray().AddNumber(0).AddNumber(0).AddNumber(12).AddNumber(12));
        PdfDictionary states = new PdfDictionary();
        states.Set("Off", file.AddObjectReference(new PdfStream(blank.Clone(), new byte[] { (byte)'q', (byte)'\n', (byte)'Q', (byte)'\n' }, false)));
        states.Set("Yes", file.AddObjectReference(new PdfStream(blank.Clone(), new byte[] { (byte)'q', (byte)'\n', (byte)'Q', (byte)'\n' }, false)));
        PdfDictionary appearance = new PdfDictionary();
        appearance.Set("N", states);
        PdfDictionary agreeDictionary = file.GetObject<PdfDictionary>(agree.Number);
        agreeDictionary.Set("AP", appearance);
        agreeDictionary.SetName("AS", "Off");
        file.MarkModified(agree.Number);
        fields.Add(agree);

        PdfDictionary acroForm = new PdfDictionary();
        acroForm.Set("Fields", fields);
        acroForm.SetString("DA", "/Helv 11 Tf 0 g");
        acroForm.SetBoolean("NeedAppearances", true);
        pageDictionary.Set("Annots", annots);
        file.MarkModified(page.Number);
        file.Catalog.Set("AcroForm", file.AddObjectReference(acroForm));
        file.MarkModified(file.CatalogNumber);
    }

    /// <summary>新建一个既是字段又是控件的终端域。Creates a terminal field that doubles as its own widget.</summary>
    /// <param name="file">目标文件。The target file.</param>
    /// <param name="page">所在页引用。The page reference.</param>
    /// <param name="name">字段名。The field name.</param>
    /// <param name="fieldType">域类型名（Tx/Ch/Btn/Sig）。The field type name.</param>
    /// <param name="flags">标志位。The flags.</param>
    /// <param name="rect">左、下、右、上四条边。The left, bottom, right and top edges.</param>
    /// <param name="annots">页面注释数组。The page annotation array.</param>
    /// <returns>新字段的引用。The reference of the new field.</returns>
    private static PdfReference AddDemoField(
        PdfFile file,
        PdfReference page,
        string name,
        string fieldType,
        int flags,
        double[] rect,
        PdfArray annots)
    {
        PdfDictionary field = new PdfDictionary();
        field.SetName("Type", "Annot");
        field.SetName("Subtype", "Widget");
        field.SetName("FT", fieldType);
        field.SetString("T", name);
        if (flags != 0) field.SetNumber("Ff", flags);
        field.Set("Rect", new PdfArray().AddNumber(rect[0]).AddNumber(rect[1]).AddNumber(rect[2]).AddNumber(rect[3]));
        field.Set("P", page);
        field.SetString("DA", "/Helv 0 Tf 0 g");
        PdfReference reference = file.AddObjectReference(field);
        annots.Add(reference);
        return reference;
    }

    /// <summary>对象级编辑演示：图层水印、页面旋转补偿、增量保存与拆页。Object-level editing demo: layered watermarks, rotation compensation, incremental save and splitting.</summary>
    /// <param name="pdfPath">上一步生成的 PDF 路径。The PDF written by the previous step.</param>
    private static void RunWatermarkDemo(string pdfPath)
    {
        string textPath = Path.Combine(OutDir, "pdf-watermark-text.pdf");
        string stampPath = Path.Combine(OutDir, "pdf-watermark-stamp.pdf");

        // 4.1 全文档背景文字水印：默认斜跨 45 度、浅灰、30% 不透明度
        PdfFile file = PdfFile.Load(pdfPath);
        file.AddWatermark("机密文件 CONFIDENTIAL");

        // 4.2 只给首页加前景平铺水印，字号、角度与不透明度自定义
        WatermarkOptions tile = WatermarkOptions.ForText("内部资料 INTERNAL USE");
        tile.Layer = WatermarkLayer.Foreground;
        tile.Placement = WatermarkPlacement.Tile;
        tile.FontSize = 28;
        tile.Bold = true;
        tile.Angle = 30;
        tile.Opacity = 0.15;
        file.AddWatermark(tile, 0);
        file.Save(textPath);

        // 4.3 先把首页转成横向阅读，再叠加右下角图像印章，最后增量保存
        PdfFile second = PdfFile.Load(textPath);
        second.SetPageRotation(0, 90);
        WatermarkOptions stamp = WatermarkOptions.ForImage(BuildStampPng());
        stamp.Placement = WatermarkPlacement.BottomRight;
        stamp.ImageWidth = 120;
        stamp.Layer = WatermarkLayer.Foreground;
        stamp.Opacity = 0.6;
        second.AddWatermark(stamp);
        second.SaveIncremental(stampPath);

        // 4.4 拆页与合并：每页单独成一份文档，再把首页与整篇合成新文档
        PdfFile source = PdfFile.Load(stampPath);
        IReadOnlyList<PdfFile> parts = source.Split(1);
        int partIndex = 0;
        foreach (PdfFile part in parts)
        {
            partIndex++;
            part.Save(Path.Combine(OutDir, "pdf-part-" + partIndex + ".pdf"));
        }
        PdfFile merged = PdfFile.Merge(parts[0], source);
        merged.Save(Path.Combine(OutDir, "pdf-merged.pdf"));

        // 4.5 直接动内容流：追加一段画线命令，再读回合并后的内容
        PdfFile annotated = PdfFile.Load(pdfPath);
        annotated.AppendPageContent(0, System.Text.Encoding.ASCII.GetBytes("q 0.9 G 20 20 m 200 20 l S Q\n"));
        int contentLength = annotated.GetPageContentBytes(0).Length;
        annotated.Save(Path.Combine(OutDir, "pdf-annotated.pdf"));

        PdfDocument stamped = PdfDocument.Load(stampPath);
        Console.WriteLine("[水印] 页数=" + stamped.PageCount + "，拆分份数=" + parts.Count
            + "，合并页数=" + merged.PageCount
            + "，追加后首页内容=" + contentLength + " 字节"
            + "，首页水印可读=" + (stamped.Pages[0].ExtractText().Contains("INTERNAL USE") ? "是" : "否")
            + "，增量文件=" + new FileInfo(stampPath).Length + " 字节");
    }

    /// <summary>生成一枚纯色矩形 PNG 作为印章图像，供图像水印使用。Builds a solid rectangle PNG used as the stamp image.</summary>
    /// <returns>PNG 字节。The PNG bytes.</returns>
    private static byte[] BuildStampPng()
    {
        int width = 96;
        int height = 48;
        byte[] raw = new byte[height * (width * 4 + 1)];
        int stride = width * 4 + 1;
        for (int y = 0; y < height; y++)
        {
            int row = y * stride;
            for (int x = 0; x < width; x++)
            {
                bool border = y < 3 || y >= height - 3 || x < 3 || x >= width - 3;
                int offset = row + 1 + x * 4;
                raw[offset] = 190;
                raw[offset + 1] = 30;
                raw[offset + 2] = 45;
                raw[offset + 3] = (byte)(border ? 255 : 40);
            }
        }

        using (MemoryStream png = new MemoryStream())
        {
            png.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, 0, 8);
            byte[] header = new byte[13];
            header[0] = (byte)(width >> 24);
            header[1] = (byte)(width >> 16);
            header[2] = (byte)(width >> 8);
            header[3] = (byte)width;
            header[4] = (byte)(height >> 24);
            header[5] = (byte)(height >> 16);
            header[6] = (byte)(height >> 8);
            header[7] = (byte)height;
            header[8] = 8;
            header[9] = 6;
            WritePngChunk(png, "IHDR", header);
            WritePngChunk(png, "IDAT", Zlib.Compress(raw));
            WritePngChunk(png, "IEND", new byte[0]);
            return png.ToArray();
        }
    }

    private static void WritePngChunk(Stream output, string type, byte[] data)
    {
        byte[] length = new byte[4];
        length[0] = (byte)(data.Length >> 24);
        length[1] = (byte)(data.Length >> 16);
        length[2] = (byte)(data.Length >> 8);
        length[3] = (byte)data.Length;
        output.Write(length, 0, 4);

        byte[] name = System.Text.Encoding.ASCII.GetBytes(type);
        output.Write(name, 0, 4);
        output.Write(data, 0, data.Length);

        uint crc = 0xFFFFFFFFu;
        for (int i = 0; i < name.Length; i++) crc = PngCrcTable[(crc ^ name[i]) & 0xFF] ^ (crc >> 8);
        for (int i = 0; i < data.Length; i++) crc = PngCrcTable[(crc ^ data[i]) & 0xFF] ^ (crc >> 8);
        crc ^= 0xFFFFFFFFu;
        byte[] tail = new byte[4];
        tail[0] = (byte)(crc >> 24);
        tail[1] = (byte)(crc >> 16);
        tail[2] = (byte)(crc >> 8);
        tail[3] = (byte)crc;
        output.Write(tail, 0, 4);
    }

    private static readonly uint[] PngCrcTable = BuildPngCrcTable();

    private static uint[] BuildPngCrcTable()
    {
        uint[] table = new uint[256];
        for (uint i = 0; i < 256; i++)
        {
            uint value = i;
            for (int bit = 0; bit < 8; bit++) value = (value & 1) != 0 ? 0xEDB88320u ^ (value >> 1) : value >> 1;
            table[i] = value;
        }
        return table;
    }

    private static int CountBlocks(OfficeDocument document)
    {
        int count = 0;
        foreach (Section section in document.Sections) count += section.Blocks.Count;
        return count;
    }

    private static string Truncate(string text, int max)
    {
        text = (text ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Trim();
        return text.Length <= max ? text : text.Substring(0, max) + "...";
    }
}
