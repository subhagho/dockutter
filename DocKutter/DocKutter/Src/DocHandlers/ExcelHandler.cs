using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using DocKutter.Common;
using DocKutter.Common.Utils;

namespace DocKutter.DocHandlers
{
    public class ExcelHandler : IDocHandler
    {
        private IDocHandlerFactory docHandlerFactory = null;
        private Application excel = null;

        public void Close()
        {
            if (excel != null)
            {
                excel.Quit();
                excel = null;
            }
        }

        public string ConvertToPDF(string fileName, string outDir, bool createDir = false)
        {
            Preconditions.CheckArgument(excel);
            Preconditions.CheckArgument(fileName);
            Preconditions.CheckArgument(outDir);

            try
            {
                lock (excel)
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

                    Workbook workbook = excel.Workbooks.Open(inFile.FullName);
                    try
                    {
                        string fname = Path.GetFileNameWithoutExtension(inFile.FullName);
                        string outpath = String.Format("{0}/{1}.PDF", outDir, fname);
                        LogUtils.Debug(String.Format("Generating PDF output. [path={0}]", outpath));

                        workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outpath, XlFixedFormatQuality.xlQualityStandard, true, true);

                        return outpath;
                    }
                    finally
                    {
                        workbook.Close();
                    }
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
            Preconditions.CheckArgument(files);
            Preconditions.CheckArgument(outDir);

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
            if (excel == null)
            {
                excel = new Application();
                excel.Visible = false;
            }
        }

        public void WithDocHandlerFactory(IDocHandlerFactory factory)
        {
            docHandlerFactory = factory;
        }
    }
}
