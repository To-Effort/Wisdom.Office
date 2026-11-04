using System;
using System.IO;
using Wisdom.Office.Core;
using Wisdom.Ppt;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// 修复规范第二集 Phase 4 的 PPT 条目：产出损坏文件与静默丢失的形状。
    /// Phase-four PPT entries of remediation design episode two: corrupt output and silently dropped shapes.
    /// </summary>
    internal static partial class Program
    {
        /// <summary>PPT Phase 4 入口。Entry point for the PPT half of phase four.</summary>
        private static void RunRemediationPptPhaseFourTest()
        {
            RunStyledPlaceholderSurvivesTest();

            Console.WriteLine("[remediation] ppt phase-four shapes and media ok");
        }

        /// <summary>
        /// PT6：库自己补出来的空占位符，被调用者加过底色或描边之后必须写得出门；
        /// 谁都没动过的那一份才允许按「PowerPoint 的空提示」省掉。
        /// PT6: a placeholder this library generated must survive the save once the caller has given it a
        /// fill or an outline; only the one nobody touched may still be dropped as PowerPoint's empty prompt.
        /// </summary>
        private static void RunStyledPlaceholderSurvivesTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();

            Slide slide = deck.AddSlide(SlideLayoutKind.TitleAndContent);
            ((TextBoxShape)slide.Shapes[0]).Paragraphs[0].AddText("标题");
            TextBoxShape block = (TextBoxShape)slide.Shapes[1];
            block.FillColor = new OfficeColor(0xFF, 0xE8, 0xD9);
            block.Outline = new BorderLine(BorderStyle.Single, 1.5, new OfficeColor(0xC0, 0x39, 0x2B));
            double width = block.Width;
            double height = block.Height;

            // 控制：另一页只填标题，版式自带的内容占位符没人动过，仍然应当被省掉——
            // 「程序生成的演示文稿不该留一个等用户点进去的空框」这条承诺不能因为本次改动丢掉。
            // Control: on the next slide only the title is filled, so the untouched layout placeholder must
            // still be dropped — the promise that generated decks carry no "click to add" boxes cannot be
            // lost with this change.
            Slide plain = deck.AddSlide(SlideLayoutKind.TitleAndContent);
            ((TextBoxShape)plain.Shapes[0]).Paragraphs[0].AddText("只有标题");
            Require(plain.Shapes.Count == 2, "the plain slide must start with two generated placeholders");

            OfficePresentation loaded;
            using (MemoryStream buffer = new MemoryStream())
            {
                PptxWriter.Write(deck, buffer);
                loaded = PptxReader.Load(buffer.ToArray());
            }

            Slide first = loaded.Slides[0];
            Require(first.Shapes.Count == 2,
                "the styled placeholder vanished on save: the slide carries " + first.Shapes.Count
                + " shape(s) instead of 2");
            TextBoxShape back = (TextBoxShape)first.Shapes[1];
            RequireHex(back.FillColor, "FFE8D9", "placeholder fill");
            Require(back.Outline != null && back.Outline.Style == BorderStyle.Single,
                "the placeholder outline was not written back");
            Require(Math.Abs(back.Width - width) < 0.5 && Math.Abs(back.Height - height) < 0.5,
                "the placeholder geometry changed: " + back.Width + " x " + back.Height);

            Slide second = loaded.Slides[1];
            Require(second.Shapes.Count == 1,
                "an untouched generated placeholder was written out: the plain slide carries "
                + second.Shapes.Count + " shape(s)");
        }
    }
}
