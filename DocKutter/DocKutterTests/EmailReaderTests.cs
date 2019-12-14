using System;
using System.Collections.Generic;
using System.IO;
using Xunit;
using DocKutter.Common.Utils;
using DocKutter.Common;

namespace DocKutter.Common
{
    public class EmailReaderTests
    {
        private static readonly string EMAIL_TEMPLATE_FILE = @"..\..\..\Resources\DemoEmailTemplate.html";
        private static readonly string EMAIL_FILE_EML = @"..\..\..\Resources\Mail\sample_email.eml";
        private static readonly string EMAIL_FILE_MSG = @"..\..\..\Resources\Mail\sample_email.msg";

        [Fact]
        public void ReadEml()
        {
            try
            {
                EmailReader reader = new EmailReader().WithHtmlMessageTemplate(EMAIL_TEMPLATE_FILE).WithDateFormat("ddd, dd MMM yyy HH:mm:ss GMT");
                string html = reader.Read(EMAIL_FILE_EML);
                Assert.True(!String.IsNullOrEmpty(html));

                string dir = FileUtils.GetTempDirectory();
                string file = String.Format("{0}/test_eml.html", dir);
                using (StreamWriter writer = new StreamWriter(file))
                {
                    writer.Write(html.ToCharArray());
                    writer.Flush();
                }
                LogUtils.Info(String.Format("Written HTML output to [{0}]", file));
            }
            catch (Exception ex)
            {
                LogUtils.Error(ex);
                throw ex;
            }
        }

        [Fact]
        public void ReadMsg()
        {
            try
            {
                EmailReader reader = new EmailReader().WithHtmlMessageTemplate(EMAIL_TEMPLATE_FILE).WithDateFormat("ddd, dd MMM yyy HH:mm:ss GMT");
                string html = reader.Read(EMAIL_FILE_MSG);
                Assert.True(!String.IsNullOrEmpty(html));

                string dir = FileUtils.GetTempDirectory();
                string file = String.Format("{0}/test_msg.html", dir);
                using (StreamWriter writer = new StreamWriter(file))
                {
                    writer.Write(html.ToCharArray());
                    writer.Flush();
                }
                LogUtils.Info(String.Format("Written HTML output to [{0}]", file));
            }
            catch (Exception ex)
            {
                LogUtils.Error(ex);
                throw ex;
            }
        }
    }
}
