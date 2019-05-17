using NPOI.OpenXml4Net.Util;
using System;
using System.Linq;
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
            return new CT_CellStyles
            {
                count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]),
                cellStyle = node.ChildNodes.Cast<XmlNode>()
                    .Where(childNode => childNode.LocalName == nameof(cellStyle))
                    .Select(childNode => CT_CellStyle.Parse(childNode, namespaceManager))
                    .ToList()
            };
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            this.cellStyle?.ForEach(x => x.Write(sw, nameof(cellStyle)));
            sw.Write($"</{nodeName}>");
        }

    }
}
