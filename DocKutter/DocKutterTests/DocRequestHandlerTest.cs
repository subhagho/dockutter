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
            LogUtils.Debug("RESPONSE", response);
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
                    string outDir = FileUtils.GetTempDirectory("EMAIL_OUTPUT");
                    foreach(string file in EMAIL_SOURCE_FILES)
                    {
                        handler.Run(DocConstants.DOC_HANDLER_EMAIL, file, outDir, new DocResponseHandler(ResponseHandler));
                    }
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
