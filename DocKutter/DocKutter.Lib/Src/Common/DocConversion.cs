using System;
using System.Collections.Generic;
using System.IO;
using LibZConfig.Common.Config.Nodes;
using DocKutter.DocHandlers;
using LibZConfig.Common.Utils;

namespace DocKutter.Common
{
    public enum ESourceType
    {
        /// <summary>
        /// Source Doc Type - Text
        /// </summary>
        TEXT,
        /// <summary>
        /// Source Doc Type - Rich Text
        /// </summary>
        RTF,
        /// <summary>
        /// Source Doc Type - Word
        /// </summary>
        DOC,
        /// <summary>
        /// Source Doc Type - Word (OpenXML)
        /// </summary>
        DOCX,
        /// <summary>
        /// Source Doc Type - Excel
        /// </summary>
        XLS,
        /// <summary>
        /// Source Doc Type - Excel (OpenXML)
        /// </summary>
        XLSX,
        /// <summary>
        /// Source Doc Type - PowerPoint
        /// </summary>
        PPT,
        /// <summary>
        /// Source Doc Type - PowerPoint (OpenXML)
        /// </summary>
        PPTX,
        /// <summary>
        /// Source Doc Type - Jpeg Image
        /// </summary>
        JPEG,
        /// <summary>
        /// Source Doc Type - Gif Image
        /// </summary>
        GIF,
        /// <summary>
        /// Source Doc Type - TIFF Image
        /// </summary>
        TIFF,
        /// <summary>
        /// Source Doc Type - PNG Image
        /// </summary>
        PNG,
        /// <summary>
        /// Source Doc Type - Html
        /// </summary>
        HTML,
        /// <summary>
        /// Source Doc Type - xHtml
        /// </summary>
        XHTML,
        /// <summary>
        /// Source Doc Type - PDF
        /// </summary>
        PDF,
        /// <summary>
        /// Source Doc Type - Email (elm file)
        /// </summary>
        ELM,
        /// <summary>
        /// Source Doc Type - Email (msg file)
        /// </summary>
        MSG
    }

    public enum ETargetType
    {
        /// <summary>
        /// Target Doc Type - PDF
        /// </summary>
        PDF,
        /// <summary>
        /// Target Doc Type - HTML
        /// </summary>
        HTML
    }

    public abstract class AbstractDocConverter : IDisposable
    {
        protected IDocHandlerFactory DocHandlerFactory { get; set; }
        protected List<ESourceType> SourceTypes { get; set; }

        public bool CanConvert(ESourceType type)
        {
            if (SourceTypes != null && SourceTypes.Count > 0)
            {
                foreach(ESourceType st in SourceTypes)
                {
                    if (st == type)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        protected static string GetOutputFile(string sourceFile, string outDir, ETargetType type, bool createDir)
        {
            FileInfo inFile = new FileInfo(sourceFile);
            if (!inFile.Exists)
            {
                throw new FileNotFoundException("Input Excel file not found.", sourceFile);
            }
            if (createDir)
            {
                if (!FileUtils.CheckDirectory(outDir))
                {
                    throw new DirectoryNotFoundException(String.Format("Output directory not found/be created. [path={0}]", outDir));
                }
            }
            string fname = Path.GetFileNameWithoutExtension(inFile.FullName);
            string outpath = String.Format("{0}/{1}.{2}", outDir, fname, type.ToString());
            LogUtils.Debug(String.Format("Generating {1} output. [path={0}]", outpath, type.ToString()));

            return outpath;
        }

        public abstract void Configure(AbstractConfigNode node);

        public abstract void Dispose();

        public abstract string Convert(string sourceFile, string ourDir, bool createDir = true, ETargetType targetFormat = ETargetType.PDF);

        public virtual Dictionary<string, string> Convert(List<string> sourceFiles, string ourDir, bool createDir = true, ETargetType targetFormat = ETargetType.PDF)
        {
            Dictionary<string, string> docSet = new Dictionary<string, string>();
            foreach(string sourceFile in sourceFiles)
            {
                string output = Convert(sourceFile, ourDir, createDir, targetFormat);
                if (!string.IsNullOrEmpty(output))
                {
                    docSet.Add(sourceFile, output);
                }
            }
            return docSet;
        }
    }
}
