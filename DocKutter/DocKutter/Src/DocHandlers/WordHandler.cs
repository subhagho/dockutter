using System;
using System.IO;
using Microsoft.Office.Interop.Word;
using DocKutter.Common;
using DocKutter.Common.Utils;

namespace DocKutter.DocHandlers
{
    public class WordHandler : IDocHandler
    {
        public string ConvertToPDF(string fileName, string outDir, bool createDir = false)
        {
            Preconditions.CheckArgument(fileName);
            try
            {
                FileInfo inFile = new FileInfo(fileName);
                if (!inFile.Exists)
                {
                    throw new FileNotFoundException("Input Excel file not found.", fileName);
                }
                if (createDir)
                {
                    if (!FileUtils.CheckDirectory(outDir))
                    {
                        throw new DirectoryNotFoundException(String.Format("Output directory not found/be created. [path={0}]", outDir));
                    }
                }
                Application word = new Application();
                word.Visible = false;
                try
                {
                    Document doc = word.Documents.Open(inFile.FullName);
                    try
                    {
                        string fname = Path.GetFileNameWithoutExtension(inFile.FullName);
                        string outpath = String.Format("{0}/{1}.PDF", outDir, fname);
                        LogUtils.Debug(String.Format("Generating PDF output. [path={0}]", outpath));

                        doc.ExportAsFixedFormat(outpath, WdExportFormat.wdExportFormatPDF);

                        return outpath;
                    }
                    finally
                    {
                        doc.Close();
                    }
                }
                finally
                {
                    word.Quit();
                }
            } 
            catch(Exception ex)
            {
                LogUtils.Error(ex);
                throw ex;
            }
        }
    }
}
