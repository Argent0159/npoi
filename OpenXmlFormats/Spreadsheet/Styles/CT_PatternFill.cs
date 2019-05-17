﻿using NPOI.OpenXml4Net.Util;
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
    public class CT_PatternFill
    {
        public bool IsSetPatternType()
        {
            return this.patternType != ST_PatternType.none;
        }
        public CT_Color AddNewFgColor()
        {
            this.fgColor = new CT_Color();
            return fgColor;
        }

        public CT_Color AddNewBgColor()
        {
            this.bgColor = new CT_Color();
            return bgColor;
        }
        public void UnsetPatternType()
        {
            this.patternType = ST_PatternType.none;
        }
        public void UnsetFgColor()
        {
            this.fgColor = null;
        }
        public void UnsetBgColor()
        {
            this.bgColor = null;
        }
        public static CT_PatternFill Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            var ctObj = new CT_PatternFill();
            if (node.Attributes[nameof(patternType)] != null)
                ctObj.patternType = (ST_PatternType)Enum.Parse(typeof(ST_PatternType), node.Attributes[nameof(patternType)].Value);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(fgColor))
                    ctObj.fgColor = CT_Color.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(bgColor))
                    ctObj.bgColor = CT_Color.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(patternType), this.patternType.ToString());
            if (this.fgColor == null && this.bgColor == null)
            {
                sw.Write("/>");
            }
            else
            {
                sw.Write(">");
                    this.fgColor?.Write(sw, nameof(fgColor));
                    this.bgColor?.Write(sw, nameof(bgColor));
                sw.Write(string.Format("</{0}>", nodeName));
            }
        }

        [XmlElement]
        public CT_Color fgColor { get; set; } = null;

        public bool IsSetBgColor()
        {
            return bgColor != null;
        }

        public bool IsSetFgColor()
        {
            return fgColor != null;
        }

        [XmlElement]
        public CT_Color bgColor { get; set; } = null;

        [XmlAttribute]
        public ST_PatternType patternType { get; set; } = ST_PatternType.none;
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_PatternType
    {


        none,


        solid,


        mediumGray,


        darkGray,


        lightGray,


        darkHorizontal,


        darkVertical,


        darkDown,


        darkUp,


        darkGrid,


        darkTrellis,


        lightHorizontal,


        lightVertical,


        lightDown,


        lightUp,


        lightGrid,


        lightTrellis,


        gray125,


        gray0625,
    }

}
