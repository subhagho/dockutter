using System;
using System.Collections.Generic;
using System.Threading;
using Xunit;
using DocKutter.Common.Utils;
using DocKutter.DocHandlers;

namespace DocKutter.Executor
{
    public class DocRequestHandlerTest
    {
        public static string[] EMAIL_SOURCE_FILES = { @"..\..\..\Resources\Mail\sample_email.eml" };


        private static void ResponseHandler(ProcessResponse response)
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

        [Fact]
        public void EmailHandlerTests()
        {
            ThreadPool.SetMaxThreads(8, 16);

            using (DocRequestHandler handler = new DocRequestHandler())
            {
                handler.Init();
                try
                {
                    List<ManualResetEvent> events = new List<ManualResetEvent>();

                    string outDir = FileUtils.GetTempDirectory("EMAIL_OUTPUT");
                    foreach(string file in EMAIL_SOURCE_FILES)
                    {
                        events.Add(handler.Run(DocConstants.DOC_HANDLER_EMAIL, file, outDir, new DocResponseHandler(ResponseHandler)));
                    }

                    WaitHandle.WaitAll(events.ToArray());
                }
                catch (Exception ex)
                {
                    LogUtils.Error(ex);
                    throw ex;
                }
            }
        }
    }
}
