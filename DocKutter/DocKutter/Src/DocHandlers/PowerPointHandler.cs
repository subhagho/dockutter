using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using DocKutter.Common;
using DocKutter.Common.Utils;

namespace DocKutter.DocHandlers
{
    public class PowerPointHandler : IDocHandler
    {
        private IDocHandlerFactory docHandlerFactory = null;
        private Application power = null;

        public void Close()
        {
            if (power != null)
            {
                power.Quit();
                power = null;
            }
        }

        public string ConvertToPDF(string fileName, string outDir, bool createDir = false)
        {
            Preconditions.CheckArgument(power);
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
            if (power == null)
            {
                power = new Application();
            }
        }

        public void WithDocHandlerFactory(IDocHandlerFactory factory)
        {
            docHandlerFactory = factory;
        }
    }
}
