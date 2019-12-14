using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using DocKutter.Common;
using DocKutter.DocHandlers;
using DocKutter.Common.Utils;

namespace DocKutter.Executor
{
    public class ProcessRequest
    {
        public IDocHandler Handler { get; set; }
        public string SourceDoc { get; set; }
        public string OutDir { get; set; }
        public DocResponseHandler Callback { get; set; }
        public ManualResetEvent Event { get; set; }
    }

    public class ProcessResponse
    {
        public string SourceDoc { get; set; }
        public string OutputFile { get; set; }
        public Exception Error { get; set; }
    }

    public delegate void DocResponseHandler(ProcessResponse response);

    public class DocRequestHandler : IDocHandlerFactory, IDisposable
    {
        
        private Dictionary<string, IDocHandler> handlers;

        public IDocHandler GetDocHandler(string name)
        {
            if (handlers.ContainsKey(name))
            {
                return handlers[name];
            }
            return null;
        }

        public void Dispose()
        {
            if (handlers != null && handlers.Count > 0)
            {
                foreach(IDocHandler handler in handlers.Values)
                {
                    handler.Close();
                }
                handlers.Clear();
            }
        }

        public void Init()
        {
            handlers = new Dictionary<string, IDocHandler>();

            handlers[DocConstants.DOC_HANDLER_EMAIL] = new OutlooklHandler();
            handlers[DocConstants.DOC_HANDLER_EMAIL].Init();
            handlers[DocConstants.DOC_HANDLER_EMAIL].WithDocHandlerFactory(this);

            handlers[DocConstants.DOC_HANDLER_EXCEL] = new ExcelHandler();
            handlers[DocConstants.DOC_HANDLER_EXCEL].Init();
            handlers[DocConstants.DOC_HANDLER_EXCEL].WithDocHandlerFactory(this);

            handlers[DocConstants.DOC_HANDLER_POWERPOINT] = new PowerPointHandler();
            handlers[DocConstants.DOC_HANDLER_POWERPOINT].Init();
            handlers[DocConstants.DOC_HANDLER_POWERPOINT].WithDocHandlerFactory(this);

            handlers[DocConstants.DOC_HANDLER_WORD] = new WordHandler();
            handlers[DocConstants.DOC_HANDLER_WORD].Init();
            handlers[DocConstants.DOC_HANDLER_WORD].WithDocHandlerFactory(this);

            handlers[DocConstants.DOC_HANDLER_HTML] = new HtmlDocHandler();
            handlers[DocConstants.DOC_HANDLER_HTML].Init();
            handlers[DocConstants.DOC_HANDLER_HTML].WithDocHandlerFactory(this);
        }

        public ManualResetEvent Run(string type, string sourceDoc, string outDir, DocResponseHandler callback)
        {
            Preconditions.CheckArgument(type);
            Preconditions.CheckArgument(sourceDoc);
            Preconditions.CheckArgument(outDir);
            Preconditions.CheckArgument(callback);

            if (handlers.ContainsKey(type))
            {
                ProcessRequest request = new ProcessRequest();
                request.Handler = handlers[type];
                request.SourceDoc = sourceDoc;
                request.OutDir = outDir;
                request.Callback = callback;
                request.Event = new ManualResetEvent(false);

                ThreadPool.QueueUserWorkItem(new WaitCallback(Process), request);

                return request.Event;
            }
            else
            {
                throw new Exception(String.Format("Document handler not found. [type={0}]", type));
            }
        }

        private void Process(object request)
        {
            ProcessRequest pr = request as ProcessRequest;

            ProcessResponse response = new ProcessResponse();
            response.SourceDoc = pr.SourceDoc;

            try
            {
                string outfile = pr.Handler.ConvertToPDF(pr.SourceDoc, pr.OutDir, true);
                LogUtils.Debug(String.Format("Generate output file. [source={0}][filename={1}]", pr.SourceDoc, outfile));
                response.OutputFile = outfile;
            }
            catch (Exception ex)
            {
                LogUtils.Error(ex);
                response.Error = ex;
            }
            finally
            {
                pr.Callback(response);
                pr.Event.Set();
            }
        }
    }
}
