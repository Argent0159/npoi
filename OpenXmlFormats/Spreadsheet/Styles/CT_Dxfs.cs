﻿using NPOI.OpenXml4Net.Util;
using System;
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
            CT_Dxfs ctObj = new CT_Dxfs();
            ctObj.count = XmlHelper.ReadUInt(node.Attributes["count"]);
            ctObj.dxf = new List<CT_Dxf>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "dxf")
                    ctObj.dxf.Add(CT_Dxf.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "count", this.count, true);
            sw.Write(">");
            if (this.dxf != null)
            {
                foreach (CT_Dxf x in this.dxf)
                {
                    x.Write(sw, "dxf");
                }
            }
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
            CT_Dxf ctObj = new CT_Dxf();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "font")
                    ctObj.font = CT_Font.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "numFmt")
                    ctObj.numFmt = CT_NumFmt.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "fill")
                    ctObj.fill = CT_Fill.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "alignment")
                    ctObj.alignment = CT_CellAlignment.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "border")
                    ctObj.border = CT_Border.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "protection")
                    ctObj.protection = CT_CellProtection.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            if (this.font != null)
                this.font.Write(sw, "font");
            if (this.numFmt != null)
                this.numFmt.Write(sw, "numFmt");
            if (this.fill != null)
                this.fill.Write(sw, "fill");
            if (this.alignment != null)
                this.alignment.Write(sw, "alignment");
            if (this.border != null)
                this.border.Write(sw, "border");
            if (this.protection != null)
                this.protection.Write(sw, "protection");
            if (this.extLst != null)
                this.extLst.Write(sw, "extLst");
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
            CT_Font font = new CT_Font();
            this.font = font;
            return font;
        }

        public CT_Fill AddNewFill()
        {
            CT_Fill fill = new CT_Fill();
            this.fill = fill;
            return fill;
        }

        public CT_Border AddNewBorder()
        {
            CT_Border border = new CT_Border();
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
