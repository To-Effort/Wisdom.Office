using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Ppt;

namespace Wisdom.Office.Tests
{
    /// <summary>Batch 4 回归：形状入场动画（p:timing）、母版主题深改（PresentationTheme 往返）与 .potx 模板格式。
    /// Batch 4 regressions: shape entrance animations (p:timing), theme deep-edit round-trip (PresentationTheme) and the .potx template format.</summary>
    internal static partial class Program
    {
        /// <summary>动画往返 + advTm 自动换页 + timing XML 结构断言。Animation round-trip plus advTm auto-advance and timing XML structure checks.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunPptAnimationTest(string outputDirectory)
        {
            OfficePresentation presentation = OfficePresentation.Widescreen();
            Slide slide = presentation.AddSlide(SlideLayoutKind.Blank);

            SlideAnimationKind[] kinds =
            {
                SlideAnimationKind.Fade, SlideAnimationKind.FlyIn, SlideAnimationKind.Wipe,
                SlideAnimationKind.Zoom, SlideAnimationKind.Split
            };
            TextBoxShape[] boxes = new TextBoxShape[kinds.Length];
            for (int i = 0; i < kinds.Length; i++)
            {
                TextBoxShape box = slide.AddTextBox(40, 40 + i * 90, 700, 70, "动画 " + kinds[i]);
                box.Animation = kinds[i];
                boxes[i] = box;
            }
            boxes[0].AnimationDelaySeconds = 0.5;
            boxes[0].AnimationDurationSeconds = 0.8;
            boxes[1].AnimationDurationSeconds = 1.5;
            slide.AdvanceAfterSeconds = 3.0;

            TextBoxShape plain = slide.AddTextBox(40, 500, 700, 30, "无动画");

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "ppt-animations.pptx"));
            PptxWriter.Write(presentation, path);
            ValidateOoxmlPackage(path, "ppt/presentation.xml");

            string slideXml = ReadArchiveEntryText(path, "ppt/slides/slide1.xml");
            Require(CountOccurrences(slideXml, "presetClass=\"entr\"") == kinds.Length,
                "entrance behaviours lost: " + CountOccurrences(slideXml, "presetClass=\"entr\""));
            Require(slideXml.IndexOf("presetId=\"10\"", StringComparison.Ordinal) >= 0, "fade preset id missing");
            Require(slideXml.IndexOf("presetId=\"2\"", StringComparison.Ordinal) >= 0, "fly-in preset id missing");
            Require(slideXml.IndexOf("presetId=\"22\"", StringComparison.Ordinal) >= 0, "wipe preset id missing");
            Require(slideXml.IndexOf("presetId=\"23\"", StringComparison.Ordinal) >= 0, "zoom preset id missing");
            Require(slideXml.IndexOf("presetId=\"26\"", StringComparison.Ordinal) >= 0, "split preset id missing");
            Require(slideXml.IndexOf("advTm=\"3000\"", StringComparison.Ordinal) >= 0, "advTm auto-advance missing");
            Require(slideXml.IndexOf("<p:timing>", StringComparison.Ordinal) >= 0, "p:timing wrapper missing");

            OfficePresentation loaded = PptxReader.Load(path);
            Slide reloaded = loaded.Slides[0];
            Require(Math.Abs(reloaded.AdvanceAfterSeconds - 3.0) < 0.01,
                "advance-after mismatch: " + reloaded.AdvanceAfterSeconds);
            TextBoxShape[] again = FindTextBoxes(reloaded);
            Require(again.Length == kinds.Length + 1, "shape count mismatch after reload: " + again.Length);
            for (int i = 0; i < kinds.Length; i++)
            {
                Require(again[i].Animation == kinds[i],
                    "animation " + i + " mismatch: " + again[i].Animation + " vs " + kinds[i]);
            }
            Require(again[0].AnimationDelaySeconds > 0.49 && again[0].AnimationDelaySeconds < 0.51,
                "delay mismatch: " + again[0].AnimationDelaySeconds);
            Require(again[0].AnimationDurationSeconds > 0.79 && again[0].AnimationDurationSeconds < 0.81,
                "duration mismatch: " + again[0].AnimationDurationSeconds);
            Require(again[1].AnimationDurationSeconds > 1.49 && again[1].AnimationDurationSeconds < 1.51,
                "fly-in duration mismatch: " + again[1].AnimationDurationSeconds);
            Require(again[2].AnimationDelaySeconds == 0.0,
                "indefinite delay read as numeric: " + again[2].AnimationDelaySeconds);
            Require(again[kinds.Length].Animation == SlideAnimationKind.None,
                "un-animated shape picked up an animation");

            Console.WriteLine("[ppt-anim] kinds=" + kinds.Length + ", advTm=3000, delay/duration round-trip ok");
        }

        /// <summary>主题深改往返：色板六阶段、主/次字体与主题名；二次写出字节稳定。Theme deep-edit round-trip: the six colour phases, major/minor fonts and the name; the second write stays byte-stable.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunPptThemeTest(string outputDirectory)
        {
            OfficePresentation presentation = OfficePresentation.Widescreen();
            Slide slide = presentation.AddSlide(SlideLayoutKind.TitleAndContent);
            // 版式占位符初始为空，写出器会跳过无文本占位符（PowerPoint 的"单击此处添加"提示），
            // 因此这里按真实用法填入标题与正文，主题字体才有 run 可落。
            // Layout placeholders start empty and the writer skips text-less ones (PowerPoint's
            // "click to add" prompt), so fill title and body the way a real deck does to give the
            // theme fonts runs to land on.
            ((TextBoxShape)slide.Shapes[0]).Paragraphs[0].AddText("主题标题");
            ((TextBoxShape)slide.Shapes[1]).Paragraphs[0].AddText("主题正文");
            presentation.Theme = new PresentationTheme
            {
                Name = "品牌主题",
                Text = new OfficeColor(0x1A, 0x1A, 0x2E),
                Background = new OfficeColor(0xFA, 0xFA, 0xF5),
                TextSecondary = new OfficeColor(0x2F, 0x4D, 0x6B),
                BackgroundSecondary = new OfficeColor(0xDD, 0xDD, 0xD2),
                Hyperlink = new OfficeColor(0xC0, 0x39, 0x2B),
                FollowedHyperlink = new OfficeColor(0x8E, 0x44, 0xAD),
                FontMajor = "微软雅黑",
                FontMinor = "Consolas"
            };

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "ppt-theme.pptx"));
            PptxWriter.Write(presentation, path);

            OfficePresentation loaded = PptxReader.Load(path);
            PresentationTheme theme = loaded.Theme;
            Require(theme != null, "theme not read back");
            Require(theme.Name == "品牌主题", "theme name mismatch: " + theme.Name);
            RequireHex(theme.Text, "1A1A2E", "text");
            RequireHex(theme.Background, "FAFAF5", "background");
            RequireHex(theme.TextSecondary, "2F4D6B", "textSecondary");
            RequireHex(theme.BackgroundSecondary, "DDDDD2", "backgroundSecondary");
            RequireHex(theme.Hyperlink, "C0392B", "hyperlink");
            RequireHex(theme.FollowedHyperlink, "8E44AD", "followedHyperlink");
            Require(theme.FontMajor == "微软雅黑", "major font mismatch: " + theme.FontMajor);
            Require(theme.FontMinor == "Consolas", "minor font mismatch: " + theme.FontMinor);

            string themeXml = ReadArchiveEntryText(path, "ppt/theme/theme1.xml");
            XDocument doc = XDocument.Parse(themeXml);
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XElement dk1 = doc.Root.Element(a + "themeElements").Element(a + "clrScheme").Element(a + "dk1");
            Require(dk1 != null && dk1.Element(a + "srgbClr") != null, "dk1 override not written as srgbClr");

            // 主题字体要落到形状文本上才算深改生效：run 的显式 a:latin 必须跟着主/次字体走。
            // Theme fonts only count as a deep edit once they reach the text: each run's explicit a:latin
            // has to follow the major/minor typefaces.
            string slideXml = ReadArchiveEntryText(path, "ppt/slides/slide1.xml");
            Require(slideXml.IndexOf("typeface=\"微软雅黑\"", StringComparison.Ordinal) >= 0,
                "theme major font not applied to title runs");
            Require(slideXml.IndexOf("typeface=\"Consolas\"", StringComparison.Ordinal) >= 0,
                "theme minor font not applied to body runs");

            // 模型显式设定的默认字体优先于主题字体。An explicit model default still outranks the theme font.
            OfficePresentation pinned = OfficePresentation.Widescreen();
            pinned.DefaultTextFormat = new TextFormat("Georgia", 16.0);
            pinned.Theme = new PresentationTheme { FontMinor = "Consolas" };
            pinned.AddSlide(SlideLayoutKind.Blank).AddTextBox(40, 40, 700, 60, "显式默认字体");
            string pinnedPath = PrepareOutputPath(Path.Combine(outputDirectory, "ppt-theme-pinned.pptx"));
            PptxWriter.Write(pinned, pinnedPath);
            string pinnedXml = ReadArchiveEntryText(pinnedPath, "ppt/slides/slide1.xml");
            Require(pinnedXml.IndexOf("typeface=\"Georgia\"", StringComparison.Ordinal) >= 0,
                "DefaultTextFormat overridden by theme font");
            Require(pinnedXml.IndexOf("typeface=\"Consolas\"", StringComparison.Ordinal) < 0,
                "theme font leaked into runs while DefaultTextFormat set");

            // 二次写出稳定：读回后再序列化，theme1.xml 与第一次完全一致。
            // Resave stability: serializing the reloaded model yields a byte-identical theme1.xml.
            string path2 = PrepareOutputPath(Path.Combine(outputDirectory, "ppt-theme-again.pptx"));
            PptxWriter.Write(loaded, path2);
            Require(ReadArchiveEntryText(path2, "ppt/theme/theme1.xml") == themeXml,
                "theme1.xml changed on resave");

            Console.WriteLine("[ppt-theme] phases=6, fonts=2, name round-trip, resave stable");
        }

        /// <summary>.potx 模板：内容类型识别、保存、加载与按扩展名写出。Potx template: content-type detection, save, load and extension-driven writing.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunPptTemplateTest(string outputDirectory)
        {
            OfficePresentation presentation = OfficePresentation.Widescreen();
            Slide slide = presentation.AddSlide(SlideLayoutKind.Title);
            TextBoxShape title = (TextBoxShape)slide.Shapes[0];
            title.Paragraphs[0].AddText("季度经营分析");
            presentation.Theme = new PresentationTheme { Name = "模板主题", FontMajor = "微软雅黑" };

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "ppt-template.potx"));
            OfficeFile source = OfficeFile.FromPresentation(presentation);
            source.Save(path);
            PptxWriter.WriteTemplate(presentation, path);   // 文件重载与门面写出等价

            Require(OfficeFileFormats.FromExtension(".potx") == OfficeFileFormat.Potx, ".potx extension mapping failed");
            Require(OfficeFileFormats.GetExtension(OfficeFileFormat.Potx) == ".potx", "potx extension mismatch");
            Require(OfficeFileFormats.GetDisplayName(OfficeFileFormat.Potx) == "PowerPoint 模板", "potx display name mismatch");
            Require(OfficeFileFormats.Detect(path) == OfficeFileFormat.Potx, "potx detection failed");

            string contentTypes = ReadArchiveEntryText(path, "[Content_Types].xml");
            Require(contentTypes.IndexOf("presentationml.template.main+xml", StringComparison.Ordinal) >= 0,
                "template content type missing");

            OfficeFile file = OfficeFile.Load(path);
            Require(file.Format == OfficeFileFormat.Potx, "facade format mismatch: " + file.Format);
            Require(file.ExtractText().IndexOf("季度经营分析", StringComparison.Ordinal) >= 0, "template text missing");
            OfficePresentation loaded = file.AsPresentation();
            Require(loaded != null && loaded.Slides.Count == 1, "template slides lost");
            Require(loaded.Theme != null && loaded.Theme.Name == "模板主题", "template theme name mismatch");

            // 普通 pptx 不能被误判为模板。A plain pptx must not be mistaken for a template.
            string pptxPath = PrepareOutputPath(Path.Combine(outputDirectory, "ppt-template-normal.pptx"));
            PptxWriter.Write(presentation, pptxPath);
            Require(OfficeFileFormats.Detect(pptxPath) == OfficeFileFormat.Pptx, "pptx misdetected as template");

            using (MemoryStream stream = new MemoryStream())
            {
                PptxWriter.WriteTemplate(presentation, stream);
                Require(stream.Length > 1024, "template stream write too small: " + stream.Length);
                Require(OfficeFileFormats.Detect(stream.ToArray()) == OfficeFileFormat.Potx,
                    "template byte detection failed");
            }

            Console.WriteLine("[ppt-potx] detect+save+load round-trip, theme preserved");
        }

        /// <summary>断言主题色读取值。Asserts one read theme colour.</summary>
        /// <param name="actual">读取到的颜色。The colour read back.</param>
        /// <param name="expectedHex">期望十六进制。The expected hex text.</param>
        /// <param name="phase">阶段名（报错用）。The phase name for errors.</param>
        private static void RequireHex(OfficeColor? actual, string expectedHex, string phase)
        {
            string hex = actual == null ? "null" : actual.Value.ToHex();
            Require(hex == expectedHex, "theme " + phase + " mismatch: " + hex);
        }

        /// <summary>按 z 序取幻灯片全部文本框。Collects every text box on the slide in z-order.</summary>
        /// <param name="slide">目标幻灯片。The target slide.</param>
        /// <returns>文本框数组。The text boxes.</returns>
        private static TextBoxShape[] FindTextBoxes(Slide slide)
        {
            System.Collections.Generic.List<TextBoxShape> boxes = new System.Collections.Generic.List<TextBoxShape>();
            for (int i = 0; i < slide.Shapes.Count; i++)
            {
                TextBoxShape box = slide.Shapes[i] as TextBoxShape;
                if (box != null) boxes.Add(box);
            }
            return boxes.ToArray();
        }

        /// <summary>读取 OOXML 包内指定条目的完整文本。Reads the full text of one entry in an OOXML package.</summary>
        /// <param name="path">包文件路径。The package path.</param>
        /// <param name="name">条目相对路径。The entry path.</param>
        /// <returns>条目文本。The entry text.</returns>
        private static string ReadArchiveEntryText(string path, string name)
        {
            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                ZipArchiveEntry entry = null;
                for (int i = 0; i < archive.Entries.Count; i++)
                {
                    if (string.Equals(archive.Entries[i].FullName.Replace('\\', '/'), name, StringComparison.OrdinalIgnoreCase))
                    {
                        entry = archive.Entries[i];
                        break;
                    }
                }
                Require(entry != null, "package entry missing: " + name);
                using (Stream stream = entry.Open())
                using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }
}
