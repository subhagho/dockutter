using System;
using System.IO;
using System.Threading;
using System.Collections.Generic;
using Xunit;
using DocKutter.Common.Utils;
using DocKutter.Executor;

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
                ThreadPool.SetMaxThreads(8, 16);

                LogUtils.Debug(String.Format("Current Directory: {0}", Directory.GetCurrentDirectory()));
                string outDir = FileUtils.GetTempDirectory(OUTPUT_DIR_NAME);

                DocRequestHandler handler = new DocRequestHandler();
                handler.Init();
                List<ManualResetEvent> events = new List<ManualResetEvent>();

                DocResponseHandler responseHandler = ResponseCallback;
                events.Add(handler.Run(DocRequestHandler.DOC_HANDLER_WORD, SOURCE_FILE, outDir, responseHandler));
                // events.Add(handler.Run(DocRequestHandler.DOC_HANDLER_WORD, SOURCE_FILE_HTML, outDir, responseHandler));
                events.Add(handler.Run(DocRequestHandler.DOC_HANDLER_WORD, SOURCE_FILE_TEMPLATE_1, outDir, responseHandler));

                WaitHandle.WaitAll(events.ToArray());
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
            }
            else
            {
                LogUtils.Info(String.Format("Processed Document. [source={0}][output={1}]", response.SourceDoc, response.OutputFile));
            }
        }
    }
}
