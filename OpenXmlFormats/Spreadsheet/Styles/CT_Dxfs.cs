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


    public class CT_Dxfs
    {
        public static CT_Dxfs Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            return new CT_Dxfs
            {
                count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]),
                dxf = node.ChildNodes.Cast<XmlNode>()
                    .Where(childNode=>childNode.LocalName == nameof(dxf))
                    .Select(childNode => CT_Dxf.Parse(childNode, namespaceManager))
                    .ToList()
            };
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count, true);
            sw.Write(">");
            this.dxf?.ForEach(x => x.Write(sw, nameof(dxf)));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Dxfs()
        {
            //this.dxfField = new List<CT_Dxf>();
        }
        [XmlElement]
        public List<CT_Dxf> dxf { get; set; }
        [XmlAttribute]
        public uint count { get; set; }

        [XmlIgnore]
        public bool countSpecified { get; set; }


    }

    public class CT_Dxf
    {
        public static CT_Dxf Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            var ctObj = new CT_Dxf();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(font))
                    ctObj.font = CT_Font.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(numFmt))
                    ctObj.numFmt = CT_NumFmt.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(fill))
                    ctObj.fill = CT_Fill.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(alignment))
                    ctObj.alignment = CT_CellAlignment.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(border))
                    ctObj.border = CT_Border.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(protection))
                    ctObj.protection = CT_CellProtection.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            this.font?.Write(sw, nameof(font));
            this.numFmt?.Write(sw, nameof(numFmt));
            this.fill?.Write(sw, nameof(fill));
            this.alignment?.Write(sw, nameof(alignment));
            this.border?.Write(sw, nameof(border));
            this.protection?.Write(sw, nameof(protection));
            this.extLst?.Write(sw, nameof(extLst));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Dxf()
        {
            //this.extLstField = new CT_ExtensionList();
            //this.protectionField = new CT_CellProtection();
            //this.borderField = new CT_Border();
            //this.alignmentField = new CT_CellAlignment();
            //this.fillField = new CT_Fill();
            //this.numFmtField = new CT_NumFmt();
            //this.fontField = new CT_Font();
        }

        public bool IsSetBorder()
        {
            return border != null;
        }
        public CT_Font AddNewFont()
        {
            var font = new CT_Font();
            this.font = font;
            return font;
        }

        public CT_Fill AddNewFill()
        {
            var fill = new CT_Fill();
            this.fill = fill;
            return fill;
        }

        public CT_Border AddNewBorder()
        {
            var border = new CT_Border();
            this.border = border;
            return border;
        }
        public bool IsSetFont()
        {
            return font != null;
        }
        public bool IsSetFill()
        {
            return fill != null;
        }
        [XmlElement]
        public CT_Font font { get; set; }
        [XmlElement]
        public CT_NumFmt numFmt { get; set; }
        [XmlElement]
        public CT_Fill fill { get; set; }
        [XmlElement]
        public CT_CellAlignment alignment { get; set; }
        [XmlElement]
        public CT_Border border { get; set; }
        [XmlElement]
        public CT_CellProtection protection { get; set; }
        [XmlElement]
        public CT_ExtensionList extLst { get; set; }
    }
}
