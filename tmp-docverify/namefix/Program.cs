using System;
using System.Text;
using Wisdom.Pdf;

namespace NameFix
{
    internal static class Program
    {
        private static void Probe(string label, string pdf)
        {
            try
            {
                PdfDocument document = PdfDocument.Load(Encoding.ASCII.GetBytes(pdf));
                Console.WriteLine(label + ": loaded, pages=" + document.PageCount);
            }
            catch (Exception ex)
            {
                Console.WriteLine(label + ": " + ex.GetType().Name + " | " + ex.Message);
            }
        }

        private static int Main()
        {
            Probe("trailer-bare-slash", "%PDF-1.7\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"
                + "2 0 obj\n<< /Type /Pages /Kids [] /Count 0 >>\nendobj\n"
                + "xref\n0 3\n0000000000 65535 f \ntrailer\n<< /Size 3 /Root 1 0 R / >>\nstartxref\n15\n%%EOF\n");

            Probe("catalog-bare-slash", "%PDF-1.7\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R / >>\nendobj\n"
                + "2 0 obj\n<< /Type /Pages /Kids [] /Count 0 >>\nendobj\n"
                + "xref\n0 3\ntrailer\n<< /Size 3 /Root 1 0 R >>\nstartxref\n15\n%%EOF\n");

            Probe("slash-then-delimiter", "%PDF-1.7\n1 0 obj\n<< /Type /Catalog /Pages ( ) >>\nendobj\n"
                + "xref\n0 2\ntrailer\n<< /Size 2 /Root 1 0 R >>\nstartxref\n9\n%%EOF\n");

            Probe("escaped-name-ok", "%PDF-1.7\n1 0 obj\n<< /Type /Catalog /PA#47ES 2 0 R >>\nendobj\n"
                + "2 0 obj\n<< /Type /Pages /Kids [] /Count 0 >>\nendobj\n"
                + "xref\n0 3\ntrailer\n<< /Size 3 /Root 1 0 R >>\nstartxref\n15\n%%EOF\n");

            Probe("clean-minimal", "%PDF-1.7\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"
                + "2 0 obj\n<< /Type /Pages /Kids [] /Count 0 >>\nendobj\n"
                + "xref\n0 3\ntrailer\n<< /Size 3 /Root 1 0 R >>\nstartxref\n15\n%%EOF\n");

            ContentStreamProbe();
            return 0;
        }

        private static void ContentStreamProbe()
        {
            string seedPath = @"D:\自研工具\office\tests\Wisdom.Office.Tests\out\facade-word.pdf";
            if (!System.IO.File.Exists(seedPath))
            {
                Console.WriteLine("content-stream: seed missing");
                return;
            }

            PdfFile file = PdfFile.Load(System.IO.File.ReadAllBytes(seedPath));
            string before = PdfDocument.Load(System.IO.File.ReadAllBytes(seedPath)).ExtractText();
            Console.WriteLine("fonts: " + file.GetPageResources(0).ToString());
            file.PrependPageContent(0, Encoding.ASCII.GetBytes("/ "));
            byte[] saved = file.SaveToBytes();
            try
            {
                PdfDocument document = PdfDocument.Load(saved);
                string text = document.ExtractText();
                Console.WriteLine("content-stream: pages=" + document.PageCount
                    + ", sameText=" + (text == before).ToString()
                    + ", len=" + text.Length + "/" + before.Length);
            }
            catch (Exception ex)
            {
                Console.WriteLine("content-stream: " + ex.GetType().Name + " | " + ex.Message);
            }
        }
    }
}
