using NPOI.OpenXml4Net.Util;
using System;
using System.Linq;
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
            return new CT_TableStyle
            {
                name = XmlHelper.ReadString(node.Attributes[nameof(name)]),
                pivot = XmlHelper.ReadBool(node.Attributes[nameof(pivot)]),
                table = XmlHelper.ReadBool(node.Attributes[nameof(table)]),
                count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]),
                tableStyleElement = node.ChildNodes.Cast<XmlNode>()
                    .Where(childNode => childNode.LocalName == nameof(tableStyleElement))
                    .Select(childNode => CT_TableStyleElement.Parse(childNode, namespaceManager))
                    .ToList()
            };
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(pivot), this.pivot);
            XmlHelper.WriteAttribute(sw, nameof(table), this.table);
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            this.tableStyleElement?.ForEach(x => x.Write(sw, nameof(tableStyleElement)));
            sw.Write($"</{nodeName}>");
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
            var ctObj = new CT_TableStyleInfo();
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            if (node.Attributes[nameof(showFirstColumn)] != null)
                ctObj.showFirstColumn = XmlHelper.ReadBool(node.Attributes[nameof(showFirstColumn)]);
            if (node.Attributes[nameof(showLastColumn)] != null)
                ctObj.showLastColumn = XmlHelper.ReadBool(node.Attributes[nameof(showLastColumn)]);
            if (node.Attributes[nameof(showRowStripes)] != null)
                ctObj.showRowStripes = XmlHelper.ReadBool(node.Attributes[nameof(showRowStripes)]);
            if (node.Attributes[nameof(showColumnStripes)] != null)
                ctObj.showColumnStripes = XmlHelper.ReadBool(node.Attributes[nameof(showColumnStripes)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(showFirstColumn), this.showFirstColumn);
            XmlHelper.WriteAttribute(sw, nameof(showLastColumn), this.showLastColumn);
            XmlHelper.WriteAttribute(sw, nameof(showRowStripes), this.showRowStripes);
            XmlHelper.WriteAttribute(sw, nameof(showColumnStripes), this.showColumnStripes);
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
            var ctObj = new CT_TableStyles();
            ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.defaultTableStyle = XmlHelper.ReadString(node.Attributes[nameof(defaultTableStyle)]);
            ctObj.defaultPivotStyle = XmlHelper.ReadString(node.Attributes[nameof(defaultPivotStyle)]);
            ctObj.tableStyle = new List<CT_TableStyle>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(tableStyle))
                    ctObj.tableStyle.Add(CT_TableStyle.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(count), this.count, true);
            XmlHelper.WriteAttribute(sw, nameof(defaultTableStyle), this.defaultTableStyle);
            XmlHelper.WriteAttribute(sw, nameof(defaultPivotStyle), this.defaultPivotStyle);
            sw.Write(">");
            this.tableStyle?.ForEach(x => x.Write(sw, nameof(tableStyle)));
            sw.Write($"</{nodeName}>");
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
            var ctObj = new CT_TableStyleElement();
            if (node.Attributes[nameof(type)] != null)
                ctObj.type = (ST_TableStyleType)Enum.Parse(typeof(ST_TableStyleType), node.Attributes[nameof(type)].Value);
            ctObj.size = XmlHelper.ReadUInt(node.Attributes[nameof(size)]);
            ctObj.dxfId = XmlHelper.ReadUInt(node.Attributes[nameof(dxfId)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(type), this.type.ToString());
            XmlHelper.WriteAttribute(sw, nameof(size), this.size);
            XmlHelper.WriteAttribute(sw, nameof(dxfId), this.dxfId);
            sw.Write(">");
            sw.Write($"</{nodeName}>");
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
