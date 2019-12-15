using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using LibZConfig.Common;
using LibZConfig.Common.Utils;
using DocKutter.Common;
using LibZConfig.Common.Config.Nodes;
using LibZConfig.Common.Config.Attributes;

namespace DocKutter.DocHandlers
{
    [ConfigPath(Path = "")]
    public class OpenXMLSpreadSheetHandler : AbstractDocConverter
    {
        public override void Configure(AbstractConfigNode node)
        {
            throw new NotImplementedException();
        }

        public override string Convert(string sourceFile, string ourDir, bool createDir = true, ETargetType targetFormat = ETargetType.PDF)
        {
            throw new NotImplementedException();
        }

        public override void Dispose()
        {
            throw new NotImplementedException();
        }
    }
}
