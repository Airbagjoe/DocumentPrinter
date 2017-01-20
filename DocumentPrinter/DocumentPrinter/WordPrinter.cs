using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using Word = Microsoft.Office.Interop.Word;

namespace DocumentPrinter
{
    class WordPrinter
    {
        private readonly Word.Application _wordInstance;

        public WordPrinter()
        {
             _wordInstance = new Word.Application();
        }

        public void PrintDocument(string filename)
        {
            var wordFile = new FileInfo(filename);
            object fileObject = wordFile.FullName;
            object oMissing = System.Reflection.Missing.Value;
            Word.Document doc = _wordInstance.Documents.Open(ref fileObject, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            doc.Activate();
            //Line spacing
            doc.Range(doc.Content.Start, doc.Content.End).Select();
            var text = _wordInstance.Selection;
            foreach (Word.Paragraph paragraph in text.Paragraphs)
            {
                paragraph.LineSpacing = 18F;
            }
            //Print
            doc.PrintOut(oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
        }

        public void Quit()
        {
            while (_wordInstance.BackgroundPrintingStatus > 0)
            {
                System.Threading.Thread.Sleep(100);
            }

            _wordInstance.Quit(false,  System.Reflection.Missing.Value,  System.Reflection.Missing.Value);
        }
    }
}
