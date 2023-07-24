using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Threading;

namespace TestDocumentFormat
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hi");
            try
            {
                using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Create("TestFile.doc", WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                    // Create the document structure and add some text.
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.WriteLine(Directory.GetFiles(Directory.GetCurrentDirectory()));
            Thread.Sleep(300000);
        }
    }
}
