using System;
using System.IO;
using Xunit;
using DocKutter.Common.Utils;

namespace DocKutter.DocHandlers
{
    public class WordHandlerTest
    {
        private const string SOURCE_FILE = @"..\..\..\Resources\Word\demo.docx";
        private const string SOURCE_FILE_HTML = @"..\..\..\Resources\Word\Working with .eml Files in MS Outlook.html";
        private const string SOURCE_FILE_TEMPLATE_1 = @"..\..\..\Resources\Word\tf89118851.dotx";
        private static string OUTPUT_DIR_NAME = System.Guid.NewGuid().ToString();
        
        [Fact]
        public void ConvertToPDF()
        {
            try
            {
                LogUtils.Debug(String.Format("Current Directory: {0}", Directory.GetCurrentDirectory()));
                string outDir = FileUtils.GetTempDirectory(OUTPUT_DIR_NAME);

                WordHandler handler = new WordHandler();
                string outfile = handler.ConvertToPDF(SOURCE_FILE, outDir, true);
            }
            catch (Exception ex)
            {
                LogUtils.Error(ex);
                throw ex;
            }
        }

        [Fact]
        public void ConvertToPDF_Html()
        {
            try
            {
                LogUtils.Debug(String.Format("Current Directory: {0}", Directory.GetCurrentDirectory()));
                string outDir = FileUtils.GetTempDirectory(OUTPUT_DIR_NAME);

                WordHandler handler = new WordHandler();
                string outfile = handler.ConvertToPDF(SOURCE_FILE_HTML, outDir, true);
            }
            catch (Exception ex)
            {
                LogUtils.Error(ex);
                throw ex;
            }
        }

        [Fact]
        public void ConvertToPDF_Template_1()
        {
            try
            {
                LogUtils.Debug(String.Format("Current Directory: {0}", Directory.GetCurrentDirectory()));
                string outDir = FileUtils.GetTempDirectory(OUTPUT_DIR_NAME);

                WordHandler handler = new WordHandler();
                string outfile = handler.ConvertToPDF(SOURCE_FILE_TEMPLATE_1, outDir, true);
            }
            catch (Exception ex)
            {
                LogUtils.Error(ex);
                throw ex;
            }
        }
    }
}
