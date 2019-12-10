using System;
using System.IO;
using Xunit;
using DocKutter.Common.Utils;

namespace DocKutter.DocHandlers
{
    public class ExcelhandlerTest
    {
        private const string SOURCE_FILE = @"..\..\..\Resources\Excel\FinancialSample.xlsx";
        private const string SOURCE_FILE_GEAR_CHARTS = @"..\..\..\Resources\Excel\GearChart.xls";
        private const string SOURCE_FILE_COUNTRY_CHARTS = @"..\..\..\Resources\Excel\CountryChart.xls";
        private static string OUTPUT_DIR_NAME = System.Guid.NewGuid().ToString();

        [Fact]
        public void ConvertToPDF()
        {
            try
            {
                LogUtils.Debug(String.Format("Current Directory: {0}", Directory.GetCurrentDirectory()));
                string outDir = FileUtils.GetTempDirectory(OUTPUT_DIR_NAME);

                ExcelHandler handler = new ExcelHandler();
                string outfile = handler.ConvertToPDF(SOURCE_FILE, outDir, true);
            }
            catch (Exception ex)
            {
                LogUtils.Error(ex);
                throw ex;
            }
        }

        [Fact]
        public void ConvertToPDF_GearCharts()
        {
            try
            {
                LogUtils.Debug(String.Format("Current Directory: {0}", Directory.GetCurrentDirectory()));
                string outDir = FileUtils.GetTempDirectory(OUTPUT_DIR_NAME);

                ExcelHandler handler = new ExcelHandler();
                string outfile = handler.ConvertToPDF(SOURCE_FILE_GEAR_CHARTS, outDir, true);
            }
            catch (Exception ex)
            {
                LogUtils.Error(ex);
                throw ex;
            }
        }

        [Fact]
        public void ConvertToPDF_CountryCharts()
        {
            try
            {
                LogUtils.Debug(String.Format("Current Directory: {0}", Directory.GetCurrentDirectory()));
                string outDir = FileUtils.GetTempDirectory(OUTPUT_DIR_NAME);

                ExcelHandler handler = new ExcelHandler();
                string outfile = handler.ConvertToPDF(SOURCE_FILE_COUNTRY_CHARTS, outDir, true);
            }
            catch (Exception ex)
            {
                LogUtils.Error(ex);
                throw ex;
            }
        }
    }
}
