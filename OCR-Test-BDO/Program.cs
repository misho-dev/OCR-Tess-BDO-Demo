using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OCR_Test_BDO
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            // Replace "image.png" with the path to your input image
            // Replace "output.docx" with the path where you want to save the output Word file
            OCRToWordExporter.ExportToWord("C:\\Users\\m.saamashvili\\source\\repos\\OCR-Test-BDO\\OCR-Test-BDO\\FAX-1.png", "output.docx");
            OCRToWordExporterPL.ExportToWord("C:\\Users\\m.saamashvili\\source\\repos\\OCR-Test-BDO\\OCR-Test-BDO\\FAX-1.png", "outputPL.docx");

            Console.WriteLine("OCR complete. Output saved to output.docx");
        }
    }
}
