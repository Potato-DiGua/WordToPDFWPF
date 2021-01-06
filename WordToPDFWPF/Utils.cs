using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace WordToPDFWPF
{
    class Utils
    {
        private static Word.Application application;
        private static Word.Application getWordApplication()
        {
            if (application == null)
            {
                application = new Word.Application
                {
                    Visible = false
                };
            }
            return application;
        }
        public static bool WordToPDF(string sourcePath)
        {
            //Console.WriteLine(sourcePath);
            bool result = false;
            Word.Document document = null;
            try
            {
                string pathWithNoExtension = sourcePath.Substring(0, sourcePath.LastIndexOf('.'));
                string PDFPath = pathWithNoExtension + ".pdf";//pdf存放位置
                if (File.Exists(sourcePath))
                {

                    for (int i = 1; File.Exists(PDFPath); i++)
                    {
                        PDFPath = pathWithNoExtension + '-' + i + ".pdf";
                    }

                    document = getWordApplication().Documents.Open(sourcePath);
                    document.ExportAsFixedFormat(PDFPath, Word.WdExportFormat.wdExportFormatPDF);
                    result = true;
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                result = false;
            }
            finally
            {
                if (document != null)
                {
                    document.Close();
                }
            }
            return result;
        }


    }
}
