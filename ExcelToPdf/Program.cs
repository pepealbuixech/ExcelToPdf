using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Linq;

namespace ExcelToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Any(arg => arg == "-h" || arg == "--help"))
            {
                GetHelp();
                return;
            }

            string path;
            if (args.Any(arg => arg == "-p"))
            {
                path = GetPath();
            }
            else
            {
                path = Directory.GetCurrentDirectory();
            }

            var rewrite = args.Any(arg => arg == "--rewrite" || arg == "-rwt");
            Convert(path, rewrite);
        }

        private static void Convert(string path, bool rewrite)
        {
            var files = Directory.GetFiles(path);
            var app = new Application();
            try
            {
                foreach (var file in files.Where(file => file.EndsWith(".xlsx")))
                {
                    var pdfName = file.Replace(".xlsx", ".pdf");

                    if (rewrite || !File.Exists(pdfName))
                    {
                        var workbook = app.Workbooks.Open(file);
                        workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdfName);
                        workbook.Close();
                    }
                }
            }
            finally
            {
                app.Quit();
            }
        }

        private static string GetPath()
        {
            
            {
                Console.WriteLine("Specify path:");
                while (true)
                {
                    var readedPath = Console.ReadLine();
                    if (Directory.Exists(readedPath))
                    {
                        return readedPath;
                    }
                    Console.WriteLine("Invalid Path");
                }
            }
        }

        private static void GetHelp()
        {
            Console.Write(@"
                This program converts the firs sheet of existing XLS in the current dir into PDF
                -> Options:
                    -rwt or --rewrite rewrite existing PDFs
                    -p specify path of directory
            ".Replace("                ", ""));
        }
    }
}
