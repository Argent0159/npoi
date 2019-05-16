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
            CT_NumFmts ctObj = new CT_NumFmts();
            ctObj.count = XmlHelper.ReadUInt(node.Attributes["count"]);
            ctObj.numFmt = new List<CT_NumFmt>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "numFmt")
                    ctObj.numFmt.Add(CT_NumFmt.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "count", this.count, true);
            sw.Write(">");
            if (this.numFmt != null)
            {
                foreach (CT_NumFmt x in this.numFmt)
                {
                    x.Write(sw, "numFmt");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_NumFmt AddNewNumFmt()
        {
            if (this.numFmt == null)
                this.numFmt = new List<CT_NumFmt>();
            CT_NumFmt newNumFmt = new CT_NumFmt();
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
