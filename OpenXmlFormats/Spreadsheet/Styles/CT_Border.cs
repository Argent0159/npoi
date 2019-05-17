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
    public enum ST_BorderStyle
    {
        none,
        thin,
        medium,
        dashed,
        dotted,
        thick,
        @double,
        hair,
        mediumDashed,
        dashDot,
        mediumDashDot,
        dashDotDot,
        mediumDashDotDot,
        slantDashDot,
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Border
    {
        public CT_Border()
        {

        }
        public override string ToString()
        {
            var ms = new MemoryStream();
            var sw = new StreamWriter(ms);
            this.Write(sw, "border");
            sw.Flush();
            ms.Position = 0;
            using (var sr = new StreamReader(ms))
            {
                return sr.ReadToEnd();
            }
        }
        public static CT_Border Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            var ctObj = new CT_Border
            {
                diagonalUp = XmlHelper.ReadBool(node.Attributes[nameof(diagonalUp)]),
                diagonalDown = XmlHelper.ReadBool(node.Attributes[nameof(diagonalDown)]),
                outline = XmlHelper.ReadBool(node.Attributes[nameof(outline)])
            };
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(left))
                    ctObj.left = CT_BorderPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(right))
                    ctObj.right = CT_BorderPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(top))
                    ctObj.top = CT_BorderPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(bottom))
                    ctObj.bottom = CT_BorderPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(diagonal))
                    ctObj.diagonal = CT_BorderPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(vertical))
                    ctObj.vertical = CT_BorderPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(horizontal))
                    ctObj.horizontal = CT_BorderPr.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(diagonalUp), this.diagonalUp, false);
            XmlHelper.WriteAttribute(sw, nameof(diagonalDown), this.diagonalDown, false);
            XmlHelper.WriteAttribute(sw, nameof(outline), this.outline, false);
            sw.Write(">");
            this.left?.Write(sw, nameof(left));
            this.right?.Write(sw, nameof(right));
            this.top?.Write(sw, nameof(top));
            this.bottom?.Write(sw, nameof(bottom));
            this.diagonal?.Write(sw, nameof(diagonal));
            this.vertical?.Write(sw, nameof(vertical));
            this.horizontal?.Write(sw, nameof(horizontal));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Border Copy()
        {
            return new CT_Border
            {
                bottom = this.bottom?.Copy(),
                top = this.top?.Copy(),
                right = this.right?.Copy(),
                left = this.left?.Copy(),
                diagonal = this.diagonal?.Copy(),
                vertical = this.vertical?.Copy(),
                horizontal = this.horizontal?.Copy(),
                diagonalUp = this.diagonalUp,
                diagonalUpSpecified = this.diagonalUpSpecified,
                diagonalDown = this.diagonalDown,
                diagonalDownSpecified = this.diagonalDownSpecified,
                outline = this.outline
            };
        }
        public CT_BorderPr AddNewDiagonal()
        {
            this.diagonal = this.diagonal ?? new CT_BorderPr();
            return this.diagonal;
        }
        public bool IsSetDiagonal()
        {
            return this.diagonal != null;
        }
        public void unsetDiagonal()
        {
            this.diagonal = null;
        }

        public void unsetRight()
        {
            this.right = null;
        }
        public void unsetLeft()
        {
            this.left = null;
        }
        public void unsetTop()
        {
            this.top = null;
        }
        public void UnsetBottom()
        {
            this.bottom = null;
        }
        public bool IsSetBottom()
        {
            return this.bottom != null;
        }
        public bool IsSetLeft()
        {
            return this.left != null;
        }
        public bool IsSetRight()
        {
            return this.right != null;
        }
        public bool IsSetTop()
        {
            return this.top != null;
        }

        public bool IsSetBorder()
        {
            return this.left != null || this.right != null
                || this.top != null || this.bottom != null;
        }
        public CT_BorderPr AddNewTop()
        {
            this.top = this.top ?? new CT_BorderPr();
            return this.top;
        }
        public CT_BorderPr AddNewRight()
        {
            this.right = this.right ?? new CT_BorderPr();
            return this.right;
        }
        public CT_BorderPr AddNewLeft()
        {
            this.left = this.left ?? new CT_BorderPr();
            return this.left;
        }
        public CT_BorderPr AddNewBottom()
        {
            this.bottom = this.bottom ?? new CT_BorderPr();
            return this.bottom;
        }

        [XmlElement(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public CT_BorderPr left { get; set; }
        [XmlElement(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public CT_BorderPr right { get; set; }
        [XmlElement(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public CT_BorderPr top { get; set; }
        [XmlElement(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public CT_BorderPr bottom { get; set; }
        [XmlElement(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public CT_BorderPr diagonal { get; set; }
        [XmlElement(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public CT_BorderPr vertical { get; set; }
        [XmlElement(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public CT_BorderPr horizontal { get; set; }
        [XmlAttribute]
        public bool diagonalUp { get; set; }

        [XmlIgnore]
        public bool diagonalUpSpecified { get; set; }
        [XmlAttribute]
        public bool diagonalDown { get; set; }

        [XmlIgnore]
        public bool diagonalDownSpecified { get; set; }
        [XmlAttribute]
        [DefaultValue(false)]
        public bool outline { get; set; }
    }
}
