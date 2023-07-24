using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace TestReportAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        [HttpGet(nameof(GetReport))]
        public ActionResult GetReport()
        {
            var data = GetFile();
            return File(data, "application/msword");
        }

        private byte[] GetFile()
        {
            using (WordprocessingDocument wordDocument =
            WordprocessingDocument.Create(@"Test.doc", WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));
            }
            return System.IO.File.ReadAllBytes("Test.doc");
        }
    }
}
