using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace DocumentPrinter.Printers
{
    public class WordPrinter : BasePrinter
    {
        public WordPrinter()
        {
             
        }

        public void PrintDocument(string filename, string printerName)
        {
            try
            {
                //Adjust to 1,5 line spacing
                var document = WordprocessingDocument.Open(filename, true);
                var paragraphs = document.MainDocumentPart.Document.Body.Descendants<Paragraph>();

                foreach (var p in paragraphs)
                {
                    // If the paragraph has no ParagraphProperties object, create one.
                    if (p.Elements<ParagraphProperties>().Count() == 0)
                    {
                        p.PrependChild(new ParagraphProperties());
                    }

                    // Get the paragraph properties element of the paragraph.
                    var pPr = p.Elements<ParagraphProperties>().First();

                    pPr.PrependChild(new SpacingBetweenLines() { Line = "360" });
                }

                document.Close();

                base.Print(filename, printerName);
            }
            catch (Exception)
            {
                ShowPrintError(filename);
                throw;
            }
           
            
        }
    }
}
