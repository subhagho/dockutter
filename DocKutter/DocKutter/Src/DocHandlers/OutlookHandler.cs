using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using DocKutter.Common;
using DocKutter.Common.Utils;

namespace DocKutter.DocHandlers
{
    public class OutlooklHandler : IDocHandler
    {
        private Application outlook = null;
        private IDocHandlerFactory docHandlerFactory = null;

        public string ConvertToPDF(string fileName, string outDir, bool createDir = false)
        {
            Preconditions.CheckArgument(outlook);
            Preconditions.CheckArgument(docHandlerFactory);
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


                string fname = Path.GetFileNameWithoutExtension(inFile.FullName);
                string ext = Path.GetExtension(inFile.FullName);
                string outpath = String.Format("{0}/{1}.PDF", outDir, fname);
                LogUtils.Debug(String.Format("Generating PDF output. [path={0}]", outpath));
                if (ext.ToLower().CompareTo("msg") == 0)
                {
                    convertMsg(outlook, inFile.FullName, outpath);
                }
                else
                {
                    convertEml(outlook, inFile.FullName, outpath);
                }
                return outpath;

            }
            catch (System.Exception ex)
            {
                LogUtils.Error(ex);
                throw ex;
            }
        }

        private void convertMsg(Application outlook, string fileName, string outfile)
        {
            MailItem mailItem = (MailItem)outlook.CreateItemFromTemplate(fileName);
            mailItem.SaveAs(outfile, OlSaveAsType.olHTML);
        }

        private void convertEml(Application outlook, string fileName, string outfile)
        {

        }

        public Dictionary<string, string> ConvertToPDF(List<string> files, string outDir, bool createDir = false)
        {
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
            if (outlook == null)
            {
                outlook = new Application();
            }
        }

        public void Close()
        {
            if (outlook != null)
            {
                outlook.Quit();
                outlook = null;
            }
        }

        public void WithDocHandlerFactory(IDocHandlerFactory factory)
        {
            docHandlerFactory = factory;
        }
    }
}
