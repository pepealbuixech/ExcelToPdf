using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Linq;

namespace ExcelToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            var app = new Application();
            var rewrite = args.Any(arg => arg == "--rewrite" || arg == "-rwt");
            var path = Directory.GetCurrentDirectory();
            var files = Directory.GetFiles(path);

            foreach (var file in files.Where(file => file.EndsWith(".xlsx")))
            {
                var workbook = app.Workbooks.Open(file);
                var pdfName = file.Replace(".xlsx", ".pdf");

                if (rewrite || !Directory.Exists(pdfName))
                {
                    workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdfName);
                }
            }

            app.Quit();
        }
    }
}
