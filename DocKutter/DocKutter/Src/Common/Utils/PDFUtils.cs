using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TheArtOfDev.HtmlRenderer.PdfSharp;

namespace DocKutter.Common.Utils
{
    public class PDFUtils
    {
        public static PdfSharp.PageSize Parse(int value)
        {
            if (Enum.IsDefined(typeof(PdfSharp.PageSize), value))
            {
                return (PdfSharp.PageSize)Enum.ToObject(typeof(PdfSharp.PageSize), value);
            }
            return PdfSharp.PageSize.A4;
        }
    }
}
