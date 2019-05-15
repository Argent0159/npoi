using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;

namespace NPOI.OpenXmlFormats.Spreadsheet.Document
{
    public class ExternalLinkDocument
    {
        public ExternalLinkDocument()
        {
        }
        public ExternalLinkDocument(CT_ExternalLink link)
        {
            this.ExternalLink = link;
        }
        public static ExternalLinkDocument Parse(XmlDocument xmldoc, XmlNamespaceManager namespaceMgr)
        {
            var obj = CT_ExternalLink.Parse(xmldoc.DocumentElement, namespaceMgr);
            return new ExternalLinkDocument(obj);
        }

        public CT_ExternalLink ExternalLink { get; set; } = null;
        public void Save(Stream stream)
        {
            using (var sw1 = new StreamWriter(stream))
            {
                this.ExternalLink.Write(sw1);
            }
        }
    }
}
