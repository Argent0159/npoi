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
            var ctObj = new CT_NumFmt();
            ctObj.numFmtId = XmlHelper.ReadUInt(node.Attributes[nameof(numFmtId)]);
            ctObj.formatCode = XmlHelper.ReadString(node.Attributes[nameof(formatCode)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
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
