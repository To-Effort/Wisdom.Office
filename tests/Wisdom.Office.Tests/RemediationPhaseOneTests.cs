using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Pdf;
using Wisdom.Ppt;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// 修复规范第二集 Phase 1 回归：备注母版形状树、页面旋转规整与字体描述符 Flags 位。
    /// Remediation design episode two, phase one: the notes-master shape tree, page-rotation
    /// normalization and the font-descriptor flag bits.
    /// </summary>
    /// <remarks>
    /// 三条断言都直接检查写出结果而不是读回后的模型，因为读取端早已把角度规整、把 Flags 原样带过，
    /// 往返断言今天就能通过，抓不到写出端的非法值。
    /// All three assertions inspect the bytes the writer emits rather than the reloaded model: the
    /// reader already normalizes rotation and carries the flags through untouched, so a round-trip
    /// assertion would pass today and hide the illegal values on the write side.
    /// </remarks>
    internal static partial class Program
    {
        /// <summary>Phase 1 入口。Entry point for phase one.</summary>
        /// <param name="pdfPath">已生成的样例 PDF。The sample PDF written earlier in the run.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunRemediationPhaseOneTest(string pdfPath, string outputDirectory)
        {
            RunNotesMasterShapeTreeTest(outputDirectory);
            RunPageRotationNormalizationTest(pdfPath, outputDirectory);
            RunFontDescriptorFlagsTest(pdfPath);

            Console.WriteLine("[remediation] phase-one notes master, page rotation and font flags ok");
        }

        /// <summary>PT1：备注母版的 p:nvGrpSpPr 必须写 p:nvPr，而不是 schema 里不存在的 p:nvGrpPr。PT1: the notes master's p:nvGrpSpPr must hold p:nvPr, never the schema-unknown p:nvGrpPr.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunNotesMasterShapeTreeTest(string outputDirectory)
        {
            OfficePresentation presentation = OfficePresentation.Widescreen();
            presentation.AddSlide(SlideLayoutKind.Blank);

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "remediation-notes-master.pptx"));
            PptxWriter.Write(presentation, path);
            ValidateOoxmlPackage(path, "ppt/presentation.xml");

            string xml = ReadArchiveEntryText(path, "ppt/notesMasters/notesMaster1.xml");
            Require(xml.IndexOf("nvGrpPr", StringComparison.Ordinal) < 0,
                "notes master writes the illegal element p:nvGrpPr");

            XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            XElement document = XElement.Parse(xml);
            XElement spTree = document.Element(p + "cSld").Element(p + "spTree");
            XElement nvGrpSpPr = spTree.Element(p + "nvGrpSpPr");
            Require(nvGrpSpPr != null, "notes master lost p:nvGrpSpPr");
            Require(nvGrpSpPr.Element(p + "cNvPr") != null, "p:nvGrpSpPr lost cNvPr");
            Require(nvGrpSpPr.Element(p + "cNvGrpSpPr") != null, "p:nvGrpSpPr lost cNvGrpSpPr");
            Require(nvGrpSpPr.Element(p + "nvPr") != null, "p:nvGrpSpPr lost the required p:nvPr child");

            OfficePresentation loaded = PptxReader.Load(path);
            Require(loaded.Slides.Count == 1, "notes-master fix broke the slide count: " + loaded.Slides.Count);
        }

        /// <summary>PD12：任意角度都要被规整成 ISO 32000-1 允许的 0/90/180/270 再写出。PD12: any angle must be normalized to the 0/90/180/270 set ISO 32000-1 allows before it is written.</summary>
        /// <param name="sourcePdfPath">样例 PDF。The sample PDF.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunPageRotationNormalizationTest(string sourcePdfPath, string outputDirectory)
        {
            byte[] source = File.ReadAllBytes(sourcePdfPath);
            int[] inputs = new int[] { 45, 91, 200, -90 };
            int[] expected = new int[] { 90, 90, 180, 270 };

            for (int i = 0; i < inputs.Length; i++)
            {
                PdfFile file = PdfFile.Load(source);
                file.SetPageRotation(0, inputs[i]);
                byte[] saved = file.SaveToBytes();
                string ascii = Encoding.GetEncoding("ISO-8859-1").GetString(saved);

                Require(ascii.IndexOf("/Rotate " + expected[i], StringComparison.Ordinal) >= 0,
                    "SetPageRotation(" + inputs[i] + ") did not write /Rotate " + expected[i]);
                Require(ascii.IndexOf("/Rotate " + inputs[i], StringComparison.Ordinal) < 0,
                    "SetPageRotation(" + inputs[i] + ") wrote a raw angle outside the 90-degree multiples ISO 32000-1 allows");

                string path = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-rotate-" + inputs[i] + ".pdf"));
                File.WriteAllBytes(path, saved);
                Require(PdfFile.Load(saved).GetPageRotation(0) == expected[i],
                    "reloaded rotation differs from the value written for input " + inputs[i]);
            }
        }

        /// <summary>PD13：字体描述符的 /Flags 按表位书写——Symbolic 恒置、Italic 只在真有斜体时置、AllCap 与 ForceBold 一类无关位不得伪造。PD13: the descriptor /Flags follows the documented bit table: Symbolic always, Italic only for a real italic face, and unrelated bits such as AllCap never forged.</summary>
        /// <param name="sourcePdfPath">样例 PDF，内含直立、粗体与斜体三种嵌入子集。The sample PDF, which embeds regular, bold and italic subsets.</param>
        private static void RunFontDescriptorFlagsTest(string sourcePdfPath)
        {
            string ascii = Encoding.GetEncoding("ISO-8859-1").GetString(File.ReadAllBytes(sourcePdfPath));
            MatchCollection matches = Regex.Matches(ascii,
                @"/FontName (\S+) /Flags (\d+) /FontBBox \[[^\]]*\] /ItalicAngle (-?[\d.]+)");
            Require(matches.Count >= 3,
                "expected at least three embedded font descriptors in the sample, found " + matches.Count);

            List<string> problems = new List<string>();
            int italicFaces = 0;
            for (int i = 0; i < matches.Count; i++)
            {
                string name = matches[i].Groups[1].Value;
                int flags = int.Parse(matches[i].Groups[2].Value);
                double italicAngle = double.Parse(matches[i].Groups[3].Value,
                    System.Globalization.CultureInfo.InvariantCulture);
                if (italicAngle != 0) italicFaces++;

                if ((flags & 4) == 0) problems.Add(name + " lost the Symbolic bit");
                if ((flags & 64) != 0) problems.Add(name + " forges the AllCap bit (/Flags " + flags + ")");
                if (italicAngle != 0 && (flags & 32) == 0) problems.Add(name + " is italic but lost the Italic bit (/Flags " + flags + ")");
                if (italicAngle == 0 && (flags & 32) != 0) problems.Add(name + " is upright but claims italic (/Flags " + flags + ")");
            }

            Require(italicFaces >= 1, "the sample lost its italic face");
            Require(problems.Count == 0, "font descriptor flags: " + string.Join(" | ", problems));
        }
    }
}
