using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Net;
using DocKutter.Common;
using DocKutter.Common.Utils;
using TheArtOfDev.HtmlRenderer.PdfSharp;
using PdfSharp.Pdf;

namespace DocKutter.DocHandlers
{
    public class HtmlDocHandler : IDocHandler
    {
        public int PageSize { get; set; }
        private IDocHandlerFactory docHandlerFactory = null;
        private WebClient client = null;

        public void Close()
        {
            client.Dispose();
        }

        public string ConvertToPDF(string fileName, string outDir, bool createDir = false)
        {
            Preconditions.CheckArgument(fileName);
            Preconditions.CheckArgument(outDir);
            try
            {
                if (createDir)
                {
                    if (!FileUtils.CheckDirectory(outDir))
                    {
                        throw new DirectoryNotFoundException(String.Format("Output directory not found/be created. [path={0}]", outDir));
                    }
                }
                Uri uri = new Uri(fileName);
                string pname = null;
                if (uri.Scheme == Uri.UriSchemeFile)
                {
                    pname = uri.LocalPath;
                    FileInfo inFile = new FileInfo(pname);
                    if (!inFile.Exists)
                    {
                        throw new FileNotFoundException("Input Excel file not found.", fileName);
                    }
                    pname = Path.GetFileNameWithoutExtension(inFile.FullName);
                }
                else if (uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeHttp)
                {
                    pname = uri.Host + "/" + uri.PathAndQuery;
                    Regex rgx = new Regex("[^a-zA-Z0-9 -]");
                    pname = rgx.Replace(pname, "_");
                }

                string outpath = String.Format("{0}/{1}.PDF", outDir, pname);
                LogUtils.Debug(String.Format("Generating PDF output. [path={0}]", outpath));

                string html = client.DownloadString(fileName);
                if (!String.IsNullOrEmpty(html))
                {
                    PdfDocument doc = PdfGenerator.GeneratePdf(html, PDFUtils.Parse(PageSize));
                    doc.Save(outpath);
                }

                return outpath;
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
            client = new WebClient();
        }

        public void WithDocHandlerFactory(IDocHandlerFactory factory)
        {
            docHandlerFactory = factory;
        }
    }
}
