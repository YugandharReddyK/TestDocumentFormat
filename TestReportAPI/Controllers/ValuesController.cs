using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Linq;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using OXmlWordProc = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using System.Reflection.Metadata.Ecma335;

namespace TestReportAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {

        [HttpGet("Test")]
        public IActionResult Test()
        {
            return Ok("Running...........");
        }

        [HttpGet(nameof(GetReport))]
        public IActionResult GetReport()
        {
            //var data = GetFile();
            try
            {
            var data = GetImagesForTemplate(@"reportData.docx");
                //return File(data, "application/msword");
                return Ok(data);

            }
            catch (System.Exception ex)
            {
                return BadRequest(ex.Message);
            }
            return null;

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

        private List<string> GetImagesForTemplate(string templateFile)
        {
            List<string> images = new List<string>();
            using (WordprocessingDocument doc = WordprocessingDocument.Open(templateFile, false))
            {
                List<OXmlWordProc.Inline> templateShapes = doc.MainDocumentPart.Document.Body
                    .Descendants<OXmlWordProc.Inline>().ToList();

                templateShapes.ForEach(templateShape =>
                {
                    string title = templateShape.DocProperties.Title;
                    if (!string.IsNullOrEmpty(title)
                        && (!title.StartsWith("SA:"))
                        && (!title.StartsWith("CHART:")))
                    {
                        images.Add(title);
                    }
                });
            }
            return images;
        }
    }
}
