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
    public enum ST_HorizontalAlignment
    {
        general,
        left,
        center,
        right,
        fill,
        justify,
        centerContinuous,
        distributed,
    }

    public enum ST_VerticalAlignment
    {
        top,
        center,
        bottom,
        justify,
        distributed,
    }

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_CellAlignment
    {

        private ST_HorizontalAlignment horizontalField;
        private ST_VerticalAlignment verticalField = ST_VerticalAlignment.bottom;
        private long textRotationField;
        private bool wrapTextField;
        private long indentField;
        private int relativeIndentField;
        private bool justifyLastLineField;
        private bool shrinkToFitField;
        private long readingOrderField;

        public static CT_CellAlignment Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            var ctObj = new CT_CellAlignment
            {
                textRotation = XmlHelper.ReadLong(node.Attributes[nameof(textRotation)]),
                wrapText = XmlHelper.ReadBool(node.Attributes[nameof(wrapText)]),
                indent = XmlHelper.ReadLong(node.Attributes[nameof(indent)]),
                relativeIndent = XmlHelper.ReadInt(node.Attributes[nameof(relativeIndent)]),
                justifyLastLine = XmlHelper.ReadBool(node.Attributes[nameof(justifyLastLine)]),
                shrinkToFit = XmlHelper.ReadBool(node.Attributes[nameof(shrinkToFit)]),
                readingOrder = XmlHelper.ReadLong(node.Attributes[nameof(readingOrder)])
            };
            if (node.Attributes[nameof(horizontal)] != null)
                ctObj.horizontal = (ST_HorizontalAlignment)Enum.Parse(typeof(ST_HorizontalAlignment), node.Attributes[nameof(horizontal)].Value);
            if (node.Attributes[nameof(vertical)] != null)
                ctObj.vertical = (ST_VerticalAlignment)Enum.Parse(typeof(ST_VerticalAlignment), node.Attributes[nameof(vertical)].Value);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            if(this.horizontal != ST_HorizontalAlignment.general)
                XmlHelper.WriteAttribute(sw, nameof(horizontal), this.horizontal.ToString());
            if (this.vertical != ST_VerticalAlignment.bottom)
                XmlHelper.WriteAttribute(sw, nameof(vertical), this.vertical.ToString());
            XmlHelper.WriteAttribute(sw, nameof(textRotation), this.textRotation);
            if(this.wrapText)
                XmlHelper.WriteAttribute(sw, nameof(wrapText), this.wrapText);
            XmlHelper.WriteAttribute(sw, nameof(indent), this.indent);
            XmlHelper.WriteAttribute(sw, nameof(relativeIndent), this.relativeIndent);
            if (justifyLastLine)
                XmlHelper.WriteAttribute(sw, nameof(justifyLastLine), this.justifyLastLine);
            if(shrinkToFit)
                XmlHelper.WriteAttribute(sw, nameof(shrinkToFit), this.shrinkToFit);
            XmlHelper.WriteAttribute(sw, nameof(readingOrder), this.readingOrder);
            sw.Write("/>");
        }

        public bool IsSetHorizontal()
        {
            return this.horizontalSpecified;
        }
        public bool IsSetVertical()
        {
            return this.verticalSpecified;
        }

        [XmlAttribute]
        [DefaultValue(ST_HorizontalAlignment.general)]
        public ST_HorizontalAlignment horizontal
        {
            get
            {
                return this.horizontalField;
            }
            set
            {
                this.horizontalField = value;
                this.horizontalSpecified = true;
            }
        }
        [XmlIgnore]
        public bool horizontalSpecified { get; set; }
        [XmlAttribute]
        [DefaultValue(ST_VerticalAlignment.bottom)]
        public ST_VerticalAlignment vertical
        {
            get
            {
                return this.verticalField;
            }
            set
            {
                this.verticalField = value;
                this.verticalSpecified = true;
            }
        }

        [XmlIgnore]
        public bool verticalSpecified { get; set; }
        [XmlAttribute]
        public long textRotation
        {
            get
            {
                return this.textRotationField;
            }
            set
            {
                this.textRotationField = value;
                this.textRotationSpecified = true;
            }
        }

        [XmlIgnore]
        public bool textRotationSpecified { get; set; }
        [XmlAttribute]
        public bool wrapText
        {
            get
            {
                return this.wrapTextField;
            }
            set
            {
                this.wrapTextField = value;
                this.wrapTextSpecified = true;
            }
        }

        [XmlIgnore]
        public bool wrapTextSpecified { get; set; }
        [XmlAttribute]
        public long indent
        {
            get
            {
                return this.indentField;
            }
            set
            {
                this.indentField = value;
                this.indentSpecified = true;
            }
        }

        [XmlIgnore]
        public bool indentSpecified { get; set; }
        [XmlAttribute]
        public int relativeIndent
        {
            get
            {
                return this.relativeIndentField;
            }
            set
            {
                this.relativeIndentField = value;
                this.relativeIndentSpecified = true;
            }
        }

        [XmlIgnore]
        public bool relativeIndentSpecified { get; set; }
        [XmlAttribute]
        public bool justifyLastLine
        {
            get
            {
                return this.justifyLastLineField;
            }
            set
            {
                this.justifyLastLineField = value;
                this.justifyLastLineSpecified = true;
            }
        }

        [XmlIgnore]
        public bool justifyLastLineSpecified { get; set; }
        [XmlAttribute]
        public bool shrinkToFit
        {
            get
            {
                return this.shrinkToFitField;
            }
            set
            {
                this.shrinkToFitField = value;
                this.shrinkToFitSpecified = true;
            }
        }

        [XmlIgnore]
        public bool shrinkToFitSpecified { get; set; }
        [XmlAttribute]
        public long readingOrder
        {
            get
            {
                return this.readingOrderField;
            }
            set
            {
                this.readingOrderField = value;
                this.readingOrderSpecified = true;
            }
        }

        [XmlIgnore]
        public bool readingOrderSpecified { get; set; }

        internal CT_CellAlignment Copy()
        {
            return new CT_CellAlignment
            {
                horizontal = this.horizontal,
                vertical = this.vertical,
                wrapText = this.wrapText,
                shrinkToFit = this.shrinkToFit,
                textRotation = this.textRotation,
                justifyLastLine = this.justifyLastLine,
                readingOrder = this.readingOrder,
                relativeIndent = this.relativeIndent,
                indent = this.indent
            };
        }
    }
}
