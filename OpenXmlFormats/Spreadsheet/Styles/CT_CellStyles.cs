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
    public class CT_CellStyles
    {
        public CT_CellStyles()
        {
            //this.cellStyleField = new List<CT_CellStyle>();
        }
        [XmlElement]
        public List<CT_CellStyle> cellStyle { get; set; }

        [XmlAttribute]
        public uint count { get; set; }

        [XmlIgnore]
        public bool countSpecified { get; set; }

        public static CT_CellStyles Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CellStyles ctObj = new CT_CellStyles();
            ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.cellStyle = new List<CT_CellStyle>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(cellStyle))
                    ctObj.cellStyle.Add(CT_CellStyle.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.cellStyle != null)
            {
                foreach (CT_CellStyle x in this.cellStyle)
                {
                    x.Write(sw, nameof(cellStyle));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

    }
}
