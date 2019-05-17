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
    public class CT_NumFmts
    {
        public CT_NumFmts()
        {
            //this.numFmtField = new List<CT_NumFmt>();
        }
        public static CT_NumFmts Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            var ctObj = new CT_NumFmts();
            ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.numFmt = new List<CT_NumFmt>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(numFmt))
                    ctObj.numFmt.Add(CT_NumFmt.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count, true);
            sw.Write(">");
            this.numFmt?.ForEach(x => x.Write(sw, nameof(numFmt)));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_NumFmt AddNewNumFmt()
        {
            this.numFmt = this.numFmt ?? new List<CT_NumFmt>();
            var newNumFmt = new CT_NumFmt();
            this.numFmt.Add(newNumFmt);
            return newNumFmt;
        }
        [XmlElement]
        public List<CT_NumFmt> numFmt { get; set; }
        [XmlAttribute]
        public uint count { get; set; }

        [XmlIgnore]
        public bool countSpecified { get; set; }
    }

}
