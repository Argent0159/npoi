using NPOI.OpenXml4Net.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class CT_TableStyle
    {
        public CT_TableStyle()
        {
            //this.tableStyleElementField = new List<CT_TableStyleElement>();
            this.pivot = true;
            this.table = true;
        }
        public static CT_TableStyle Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TableStyle ctObj = new CT_TableStyle();
            ctObj.name = XmlHelper.ReadString(node.Attributes["name"]);
            ctObj.pivot = XmlHelper.ReadBool(node.Attributes["pivot"]);
            ctObj.table = XmlHelper.ReadBool(node.Attributes["table"]);
            ctObj.count = XmlHelper.ReadUInt(node.Attributes["count"]);
            ctObj.tableStyleElement = new List<CT_TableStyleElement>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "tableStyleElement")
                    ctObj.tableStyleElement.Add(CT_TableStyleElement.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "name", this.name);
            XmlHelper.WriteAttribute(sw, "pivot", this.pivot);
            XmlHelper.WriteAttribute(sw, "table", this.table);
            XmlHelper.WriteAttribute(sw, "count", this.count);
            sw.Write(">");
            if (this.tableStyleElement != null)
            {
                foreach (CT_TableStyleElement x in this.tableStyleElement)
                {
                    x.Write(sw, "tableStyleElement");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [XmlElement]
        public List<CT_TableStyleElement> tableStyleElement { get; set; }
        [XmlAttribute]
        public string name { get; set; }
        [XmlAttribute]
        [DefaultValue(true)]
        public bool pivot { get; set; }
        [XmlAttribute]
        [DefaultValue(true)]
        public bool table { get; set; }
        [XmlAttribute]
        public uint count { get; set; }

        [XmlIgnore]
        public bool countSpecified { get; set; }
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_TableStyleInfo
    {
        public static CT_TableStyleInfo Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TableStyleInfo ctObj = new CT_TableStyleInfo();
            ctObj.name = XmlHelper.ReadString(node.Attributes["name"]);
            if (node.Attributes["showFirstColumn"] != null)
                ctObj.showFirstColumn = XmlHelper.ReadBool(node.Attributes["showFirstColumn"]);
            if (node.Attributes["showLastColumn"] != null)
                ctObj.showLastColumn = XmlHelper.ReadBool(node.Attributes["showLastColumn"]);
            if (node.Attributes["showRowStripes"] != null)
                ctObj.showRowStripes = XmlHelper.ReadBool(node.Attributes["showRowStripes"]);
            if (node.Attributes["showColumnStripes"] != null)
                ctObj.showColumnStripes = XmlHelper.ReadBool(node.Attributes["showColumnStripes"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "name", this.name);
            XmlHelper.WriteAttribute(sw, "showFirstColumn", this.showFirstColumn);
            XmlHelper.WriteAttribute(sw, "showLastColumn", this.showLastColumn);
            XmlHelper.WriteAttribute(sw, "showRowStripes", this.showRowStripes);
            XmlHelper.WriteAttribute(sw, "showColumnStripes", this.showColumnStripes);
            sw.Write("/>");
        }

        [XmlAttribute]
        public string name { get; set; }
        [XmlAttribute]
        public bool showFirstColumn { get; set; }

        [XmlIgnore]
        public bool showFirstColumnSpecified { get; set; }
        [XmlAttribute]
        public bool showLastColumn { get; set; }

        [XmlIgnore]
        public bool showLastColumnSpecified { get; set; }
        [XmlAttribute]
        public bool showRowStripes { get; set; }

        [XmlIgnore]
        public bool showRowStripesSpecified { get; set; }
        [XmlAttribute]
        public bool showColumnStripes { get; set; }

        [XmlIgnore]
        public bool showColumnStripesSpecified { get; set; }
    }
    public class CT_TableStyles
    {
        public static CT_TableStyles Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TableStyles ctObj = new CT_TableStyles();
            ctObj.count = XmlHelper.ReadUInt(node.Attributes["count"]);
            ctObj.defaultTableStyle = XmlHelper.ReadString(node.Attributes["defaultTableStyle"]);
            ctObj.defaultPivotStyle = XmlHelper.ReadString(node.Attributes["defaultPivotStyle"]);
            ctObj.tableStyle = new List<CT_TableStyle>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "tableStyle")
                    ctObj.tableStyle.Add(CT_TableStyle.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "count", this.count, true);
            XmlHelper.WriteAttribute(sw, "defaultTableStyle", this.defaultTableStyle);
            XmlHelper.WriteAttribute(sw, "defaultPivotStyle", this.defaultPivotStyle);
            sw.Write(">");
            if (this.tableStyle != null)
            {
                foreach (CT_TableStyle x in this.tableStyle)
                {
                    x.Write(sw, "tableStyle");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_TableStyles()
        {
            //this.tableStyleField = new List<CT_TableStyle>();
        }
        [XmlElement]
        public List<CT_TableStyle> tableStyle { get; set; }
        [XmlAttribute]
        public uint count { get; set; }

        [XmlIgnore]
        public bool countSpecified { get; set; }

        public string defaultTableStyle { get; set; }

        public string defaultPivotStyle { get; set; }
    }
    public enum ST_TableStyleType
    {


        wholeTable,


        headerRow,


        totalRow,


        firstColumn,


        lastColumn,


        firstRowStripe,


        secondRowStripe,


        firstColumnStripe,


        secondColumnStripe,


        firstHeaderCell,


        lastHeaderCell,


        firstTotalCell,


        lastTotalCell,


        firstSubtotalColumn,


        secondSubtotalColumn,


        thirdSubtotalColumn,


        firstSubtotalRow,


        secondSubtotalRow,


        thirdSubtotalRow,


        blankRow,


        firstColumnSubheading,


        secondColumnSubheading,


        thirdColumnSubheading,


        firstRowSubheading,


        secondRowSubheading,


        thirdRowSubheading,


        pageFieldLabels,


        pageFieldValues,
    }

    public class CT_TableStyleElement
    {
        public CT_TableStyleElement()
        {
            this.size = (uint)(1);
        }
        public static CT_TableStyleElement Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TableStyleElement ctObj = new CT_TableStyleElement();
            if (node.Attributes["type"] != null)
                ctObj.type = (ST_TableStyleType)Enum.Parse(typeof(ST_TableStyleType), node.Attributes["type"].Value);
            ctObj.size = XmlHelper.ReadUInt(node.Attributes["size"]);
            ctObj.dxfId = XmlHelper.ReadUInt(node.Attributes["dxfId"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "type", this.type.ToString());
            XmlHelper.WriteAttribute(sw, "size", this.size);
            XmlHelper.WriteAttribute(sw, "dxfId", this.dxfId);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [XmlAttribute]
        public ST_TableStyleType type { get; set; }

        [DefaultValue(typeof(uint), "1")]
        public uint size { get; set; }
        [XmlAttribute]
        public uint dxfId { get; set; }

        [XmlIgnore]
        public bool dxfIdSpecified { get; set; }
    }
}
