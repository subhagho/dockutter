using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MsgReader;
using DocKutter.Common.Utils;
using System.Globalization;
using HtmlAgilityPack;

namespace DocKutter.Common
{
    public class EmailReader
    {
        public static readonly string FIELD_REGEX = @"\$\{([A-Za-z,;:'\s<>=@&/]+)::\s*(__\w+)\s*\}";
        public static readonly string FIELD_FROM = "__FROM";
        public static readonly string FIELD_TO = "__TO";
        public static readonly string FIELD_CC = "__CC";
        public static readonly string FIELD_BCC = "__BCC";
        public static readonly string FIELD_SUBJECT = "__SUBJECT";
        public static readonly string FIELD_RECEIVED = "__RECEIVED";
        public static readonly string FIELD_SENT = "__SENT";
        public static readonly string FIELD_BODY = "__BODY";
        public static readonly string NODE_HEADER = "__EMAIL_HEADER";

        private string htmlMessageTemplate = null;
        private string dateFormat = CultureInfo.CurrentCulture.DateTimeFormat.LongDatePattern;

        public EmailReader WithDateFormat(string dateFormat)
        {
            Preconditions.CheckArgument(dateFormat);
            this.dateFormat = dateFormat;

            return this;
        }

        public EmailReader WithHtmlMessageTemplate(string htmlMessageTemplateFile)
        {
            Preconditions.CheckArgument(htmlMessageTemplateFile);
            FileInfo fi = new FileInfo(htmlMessageTemplateFile);
            if (!fi.Exists)
            {
                throw new FileNotFoundException("HTML Template file not found.", htmlMessageTemplateFile);
            }
            using (StreamReader reader = new StreamReader(fi.FullName))
            {
                StringBuilder builder = new StringBuilder();
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    builder.Append(line);
                }
                htmlMessageTemplate = builder.ToString();
            }
            return this;
        }

        public string Read(string file)
        {
            Preconditions.CheckArgument(file);
            Preconditions.CheckArgument(htmlMessageTemplate);

            FileInfo inFile = new FileInfo(file);
            if (!inFile.Exists)
            {
                throw new FileNotFoundException("Input Excel file not found.", file);
            }
            string ext = Path.GetExtension(inFile.FullName);

            if (ext.ToLower().CompareTo(".msg") == 0)
            {
                MsgReader.Outlook.Storage.Message message = new MsgReader.Outlook.Storage.Message(file);
                MatchCollection matches = Regex.Matches(htmlMessageTemplate, FIELD_REGEX);
                if (matches != null && matches.Count > 0)
                {
                    string html = new string(htmlMessageTemplate.ToCharArray());
                    string htmlBody = null;
                    foreach (Match match in matches)
                    {
                        string k = match.Groups[1].Value;
                        string v = match.Groups[2].Value;
                        string t = null;
                        if (v.CompareTo(FIELD_BCC) == 0)
                        {
                            t = message.GetEmailRecipients(MsgReader.Outlook.RecipientType.Bcc, true, true);
                        }
                        else if (v.CompareTo(FIELD_CC) == 0)
                        {
                            t = message.GetEmailRecipients(MsgReader.Outlook.RecipientType.Cc, true, true);
                        }
                        else if (v.CompareTo(FIELD_TO) == 0)
                        {
                            t = message.GetEmailRecipients(MsgReader.Outlook.RecipientType.To, true, true);
                        }
                        else if (v.CompareTo(FIELD_FROM) == 0)
                        {
                            t = message.GetEmailSender(true, true);
                        }
                        else if (v.CompareTo(FIELD_RECEIVED) == 0)
                        {
                            DateTime dt = (DateTime)message.ReceivedOn;
                            t = dt.ToString(dateFormat);
                        }
                        else if (v.CompareTo(FIELD_SENT) == 0)
                        {
                            DateTime dt = (DateTime)message.SentOn;
                            t = dt.ToString(dateFormat);
                        }
                        else if (v.CompareTo(FIELD_SUBJECT) == 0)
                        {
                            t = message.Subject;
                        }
                        else if (v.CompareTo(FIELD_BODY) == 0)
                        {
                            if (message.BodyHtml != null)
                            {
                                htmlBody = message.BodyHtml;
                                continue;
                            }
                            else
                            {
                                t = message.BodyText;
                            }
                        }

                        if (!String.IsNullOrEmpty(t))
                        {
                            t = String.Format("{0}{1}", k, t);
                            html = ReplaceMatch(html, match, t);
                        }
                    }
                    return PostProcess(html, htmlBody);
                }
            }
            else
            {
                MsgReader.Mime.Message message = MsgReader.Mime.Message.Load(inFile);
                MatchCollection matches = Regex.Matches(htmlMessageTemplate, FIELD_REGEX);
                if (matches != null && matches.Count > 0)
                {
                    string html = new string(htmlMessageTemplate.ToCharArray());
                    string htmlBody = null;
                    foreach (Match match in matches)
                    {
                        string k = match.Groups[1].Value;
                        string v = match.Groups[2].Value;
                        string t = null;
                        if (v.CompareTo(FIELD_BCC) == 0)
                        {
                            t = GetRecepientsString(message.Headers.Bcc);
                        }
                        else if (v.CompareTo(FIELD_CC) == 0)
                        {
                            t = GetRecepientsString(message.Headers.Cc);
                        }
                        else if (v.CompareTo(FIELD_TO) == 0)
                        {
                            t = GetRecepientsString(message.Headers.To);
                        }
                        else if (v.CompareTo(FIELD_FROM) == 0)
                        {
                            t = EmailAddressString(message.Headers.From);
                        }
                        else if (v.CompareTo(FIELD_RECEIVED) == 0)
                        {
                            // TODO: Need to fix this.
                            DateTime dt = DateTime.Now;
                            t = dt.ToString(dateFormat);
                        }
                        else if (v.CompareTo(FIELD_SENT) == 0)
                        {
                            DateTime dt = (DateTime)message.Headers.DateSent;
                            t = dt.ToString(dateFormat);
                        }
                        else if (v.CompareTo(FIELD_SUBJECT) == 0)
                        {
                            t = message.Headers.Subject;
                        }
                        else if (v.CompareTo(FIELD_BODY) == 0)
                        {
                            if (message.HtmlBody != null)
                            {
                                htmlBody = message.HtmlBody.GetBodyAsText();
                                continue;
                            }
                            else
                            {
                                t = message.TextBody.GetBodyAsText();
                            }
                        }

                        if (!String.IsNullOrEmpty(t))
                        {
                            t = String.Format("{0}{1}", k, t);
                            html = ReplaceMatch(html, match, t);
                        }
                    }
                    return PostProcess(html, htmlBody);
                }
            }
            return null;
        }

        private string PostProcess(string html, string htmlBody)
        {
            html = Regex.Replace(html, FIELD_REGEX, "");
            if (!String.IsNullOrEmpty(htmlBody))
            {
                HtmlDocument body = new HtmlDocument();
                body.LoadHtml(htmlBody);

                HtmlDocument header = new HtmlDocument();
                header.LoadHtml(html);

                var hnode = header.GetElementbyId(NODE_HEADER);
                if (hnode == null)
                {
                    throw new NodeNotFoundException(String.Format("Header node not found. [id={0}]", NODE_HEADER));
                }
                var bnode = body.DocumentNode.SelectSingleNode("//body");
                if (bnode == null)
                {
                    throw new NodeNotFoundException(string.Format("Body node not found. [name={0}]", "body"));
                }
                var nhnode = hnode.CloneNode(true);
                bnode.InsertBefore(nhnode, bnode.ChildNodes[0]);

                return body.DocumentNode.InnerHtml;
            }
            return html;
        }

        private string GetRecepientsString(List<MsgReader.Mime.Header.RfcMailAddress> addresses)
        {
            StringBuilder builder = new StringBuilder();
            foreach (MsgReader.Mime.Header.RfcMailAddress address in addresses)
            {
                if (builder.Length > 0)
                {
                    builder.Append(",");
                }
                builder.Append(EmailAddressString(address));
            }
            return builder.ToString();
        }

        private string EmailAddressString(MsgReader.Mime.Header.RfcMailAddress address, bool href = true)
        {
            string str = String.Format("{0}, {1}", address.DisplayName, address.Address);
            if (href)
            {
                str = String.Format("<a href='mailto:{0}'>{1}</a>", address.Address, str);
            }
            return str;
        }

        private string ReplaceMatch(string source, Match match, string value)
        {
            string m = match.Groups[0].Value;
            source = source.Replace(m, value);
            return source;
        }
    }
}
