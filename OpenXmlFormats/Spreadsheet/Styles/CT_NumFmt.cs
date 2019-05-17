using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_NumFmt
    {
        public static CT_NumFmt Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            return new CT_NumFmt
            {
                numFmtId = XmlHelper.ReadUInt(node.Attributes[nameof(numFmtId)]),
                formatCode = XmlHelper.ReadString(node.Attributes[nameof(formatCode)])
            };
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(numFmtId), this.numFmtId,true);
            XmlHelper.WriteAttribute(sw, nameof(formatCode), this.formatCode);
            sw.Write("/>");
        }

        [XmlAttribute]
        public uint numFmtId { get; set; }
        [XmlAttribute]
        public string formatCode { get; set; }
    }
}
