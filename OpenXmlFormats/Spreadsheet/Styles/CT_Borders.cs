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
    public class CT_Borders
    {
        public static CT_Borders Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Borders ctObj = new CT_Borders();
            ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.border = new List<CT_Border>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(border))
                    ctObj.border.Add(CT_Border.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count, true);
            sw.Write(">");
            this.border?.ForEach(x => x.Write(sw, nameof(border)));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Borders()
        {
            //this.borderField = new List<CT_Border>();
        }
        public CT_Border AddNewBorder()
        {
            this.border = this.border ?? new List<CT_Border>();
            CT_Border border = new CT_Border();
            this.border.Add(border);
            return border;
        }
        [XmlElement]
        public List<CT_Border> border { get; set; }
        public void SetBorderArray(List<CT_Border> array)
        {
            border = array;
        }
        [XmlAttribute]
        public uint count { get; set; } = 0;

        [XmlIgnore]
        public bool countSpecified { get; set; } = false;
    }
}
