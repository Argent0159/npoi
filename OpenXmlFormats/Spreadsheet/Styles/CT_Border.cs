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
            MemoryStream ms = new MemoryStream();
            StreamWriter sw = new StreamWriter(ms);
            this.Write(sw, "border");
            sw.Flush();
            ms.Position = 0;
            using (StreamReader sr = new StreamReader(ms))
            {
                return sr.ReadToEnd();
            }
        }
        public static CT_Border Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Border ctObj = new CT_Border();
            ctObj.diagonalUp = XmlHelper.ReadBool(node.Attributes[nameof(diagonalUp)]);
            ctObj.diagonalDown = XmlHelper.ReadBool(node.Attributes[nameof(diagonalDown)]);
            ctObj.outline = XmlHelper.ReadBool(node.Attributes[nameof(outline)]);
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
            if (this.left != null)
                this.left.Write(sw, nameof(left));
            if (this.right != null)
                this.right.Write(sw, nameof(right));
            if (this.top != null)
                this.top.Write(sw, nameof(top));
            if (this.bottom != null)
                this.bottom.Write(sw, nameof(bottom));
            if (this.diagonal != null)
                this.diagonal.Write(sw, nameof(diagonal));
            if (this.vertical != null)
                this.vertical.Write(sw, nameof(vertical));
            if (this.horizontal != null)
                this.horizontal.Write(sw, nameof(horizontal));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Border Copy()
        {
            CT_Border obj = new CT_Border();
            obj.bottom = this.bottom == null ? null : this.bottom.Copy();
            obj.top = this.top == null ? null : this.top.Copy();
            obj.right = this.right == null ? null : this.right.Copy();
            obj.left = this.left == null ? null : this.left.Copy();

            obj.diagonal = this.diagonal == null ? null : this.diagonal.Copy();
            obj.vertical = this.vertical == null ? null : this.vertical.Copy();
            obj.horizontal = this.horizontal == null ? null : this.horizontal.Copy();

            obj.diagonalUp = this.diagonalUp;
            obj.diagonalUpSpecified = this.diagonalUpSpecified;
            obj.diagonalDown = this.diagonalDown;
            obj.diagonalDownSpecified = this.diagonalDownSpecified;
            obj.outline = this.outline;
            return obj;
        }
        public CT_BorderPr AddNewDiagonal()
        {
            if (this.diagonal == null)
                this.diagonal = new CT_BorderPr();
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
            if (this.top == null)
                this.top = new CT_BorderPr();
            return this.top;
        }
        public CT_BorderPr AddNewRight()
        {
            if (this.right == null)
                this.right = new CT_BorderPr();
            return this.right;
        }
        public CT_BorderPr AddNewLeft()
        {
            if (this.left == null)
                this.left = new CT_BorderPr();
            return this.left;
        }
        public CT_BorderPr AddNewBottom()
        {
            if (this.bottom == null)
                this.bottom = new CT_BorderPr();
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
