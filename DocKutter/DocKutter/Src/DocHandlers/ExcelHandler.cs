using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using DocKutter.Common;
using DocKutter.Common.Utils;

namespace DocKutter.DocHandlers
{
    public class ExcelHandler : IDocHandler
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
                Application excel = new Application();
                excel.Visible = false;
                try
                {
                    Workbook workbook = excel.Workbooks.Open(inFile.FullName);
                    try
                    {
                        string fname = Path.GetFileNameWithoutExtension(fileName);
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
                finally
                {
                    excel.Quit();
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
