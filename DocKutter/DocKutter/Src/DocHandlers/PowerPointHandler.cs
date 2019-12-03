using System;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using DocKutter.Common;
using DocKutter.Common.Utils;

namespace DocKutter.DocHandlers
{
    public class PowerPointHandler : IDocHandler
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
                Application power = new Application();
               
                try
                {
                    Presentation pp = power.Presentations.Open(inFile.FullName);
                    try
                    {
                        string fname = Path.GetFileNameWithoutExtension(inFile.FullName);
                        string outpath = String.Format("{0}/{1}.PDF", outDir, fname);
                        LogUtils.Debug(String.Format("Generating PDF output. [path={0}]", outpath));

                        pp.ExportAsFixedFormat(outpath, PpFixedFormatType.ppFixedFormatTypePDF, PpFixedFormatIntent.ppFixedFormatIntentPrint);

                        return outpath;
                    }
                    finally
                    {
                        pp.Close();
                    }
                }
                finally
                {
                    power.Quit();
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
