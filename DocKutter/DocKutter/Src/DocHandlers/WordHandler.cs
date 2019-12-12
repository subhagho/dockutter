using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using DocKutter.Common;
using DocKutter.Common.Utils;

namespace DocKutter.DocHandlers
{
    public class WordHandler : IDocHandler
    {
        private Application word = null;

        public void Close()
        {
            if (word != null)
            {
                word.Quit();
                word = null;
            }
        }

        public string ConvertToPDF(string fileName, string outDir, bool createDir = false)
        {
            Preconditions.CheckArgument(word);
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
            catch (Exception ex)
            {
                LogUtils.Error(ex);
                throw ex;
            }
        }

        public Dictionary<string, string> ConvertToPDF(List<string> files, string outDir, bool createDir = false)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            foreach (string file in files)
            {
                string output = ConvertToPDF(file, outDir, createDir);
                result.Add(file, output);
            }
            return result;
        }

        public void Init()
        {
            if (word == null)
            {
                word = new Application();
                word.Visible = false;
            }
        }
    }
}
