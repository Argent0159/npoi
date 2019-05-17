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
    public class CT_Fills
    {
        public CT_Fills()
        {
            //this.fillField = new List<CT_Fill>();
        }
        public static CT_Fills Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            return new CT_Fills
            {
                count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]),
                fill = node.ChildNodes.Cast<XmlNode>()
                    .Where(childNode => childNode.LocalName == nameof(fill))
                    .Select(childNode => CT_Fill.Parse(childNode, namespaceManager))
                    .ToList()
            };
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            this.fill?.ForEach(x => x.Write(sw, nameof(fill)));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [XmlElement]
        public List<CT_Fill> fill { get; set; }
        public void SetFillArray(List<CT_Fill> array)
        {
            fill = array;
        }
        [XmlAttribute]
        public uint count { get; set; } = 0;

        [XmlIgnore]
        public bool countSpecified { get; set; } = false;
    }
}
