using System;
using System.IO;
using System.Threading;
using System.Collections.Generic;
using Xunit;
using DocKutter.Common.Utils;
using DocKutter.Executor;

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
                ThreadPool.SetMaxThreads(8, 16);

                LogUtils.Debug(String.Format("Current Directory: {0}", Directory.GetCurrentDirectory()));
                string outDir = FileUtils.GetTempDirectory(OUTPUT_DIR_NAME);

                using (DocRequestHandler handler = new DocRequestHandler())
                {
                    handler.Init();
                    List<ManualResetEvent> events = new List<ManualResetEvent>();

                    DocResponseHandler responseHandler = ResponseCallback;
                    events.Add(handler.Run(DocConstants.DOC_HANDLER_EXCEL, SOURCE_FILE, outDir, responseHandler));
                    events.Add(handler.Run(DocConstants.DOC_HANDLER_EXCEL, SOURCE_FILE_COUNTRY_CHARTS, outDir, responseHandler));
                    events.Add(handler.Run(DocConstants.DOC_HANDLER_EXCEL, SOURCE_FILE_GEAR_CHARTS, outDir, responseHandler));

                    WaitHandle.WaitAll(events.ToArray());
                }
            }
            catch (Exception ex)
            {
                LogUtils.Error(ex);
                throw ex;
            }
        }

        private void ResponseCallback(ProcessResponse response)
        {
            if (response.Error != null)
            {
                LogUtils.Error(response.Error);
                throw response.Error;
            }
            else
            {
                LogUtils.Info(String.Format("Processed Document. [source={0}][output={1}]", response.SourceDoc, response.OutputFile));
            }
        }
    }
}
