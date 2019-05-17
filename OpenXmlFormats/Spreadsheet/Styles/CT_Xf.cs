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
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(ElementName = "xf", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public class CT_Xf
    {
        private uint numFmtIdField = 0;
        private uint fontIdField = 0;
        private uint fillIdField = 0;
        private uint borderIdField = 0;
        private uint xfIdField = 0;

        public CT_Xf Copy()
        {
            CT_Xf obj = new CT_Xf();
            obj.alignment = this.alignment?.Copy();
            obj.protection = this.protection;
            obj.extLst = this.extLst?.Copy();
            obj.applyAlignment = this.applyAlignment;
            obj.applyBorder = this.applyBorder;
            obj.applyFill = this.applyFill;
            obj.applyFont = this.applyFont;
            obj.applyNumberFormat = this.applyNumberFormat;
            obj.applyProtection = this.applyProtection;
            obj.borderId = this.borderId;
            obj.borderIdSpecified = this.borderIdSpecified;
            obj.fillId = this.fillId;
            obj.fillIdSpecified = this.fillIdSpecified;
            obj.fontId = this.fontId;
            obj.fontIdSpecified = this.fontIdSpecified;
            obj.numFmtId = this.numFmtId;
            obj.numFmtIdSpecified = this.numFmtIdSpecified;
            obj.pivotButton = this.pivotButton;
            obj.quotePrefix = this.quotePrefix;
            obj.xfIdField = this.xfIdField;
            return obj;
        }

        public static CT_Xf Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Xf ctObj = new CT_Xf();
            ctObj.numFmtId = XmlHelper.ReadUInt(node.Attributes[nameof(numFmtId)]);
            ctObj.fontId = XmlHelper.ReadUInt(node.Attributes[nameof(fontId)]);
            ctObj.fillId = XmlHelper.ReadUInt(node.Attributes[nameof(fillId)]);
            ctObj.borderId = XmlHelper.ReadUInt(node.Attributes[nameof(borderId)]);
            ctObj.xfId = XmlHelper.ReadUInt(node.Attributes[nameof(xfId)]);
            ctObj.quotePrefix = XmlHelper.ReadBool(node.Attributes[nameof(quotePrefix)]);
            ctObj.pivotButton = XmlHelper.ReadBool(node.Attributes[nameof(pivotButton)]);
            ctObj.applyNumberFormat = XmlHelper.ReadBool(node.Attributes[nameof(applyNumberFormat)]);
            ctObj.applyFont = XmlHelper.ReadBool(node.Attributes[nameof(applyFont)]);
            ctObj.applyFill = XmlHelper.ReadBool(node.Attributes[nameof(applyFill)]);
            ctObj.applyBorder = XmlHelper.ReadBool(node.Attributes[nameof(applyBorder)]);
            ctObj.applyAlignment = XmlHelper.ReadBool(node.Attributes[nameof(applyAlignment)]);
            ctObj.applyProtection = XmlHelper.ReadBool(node.Attributes[nameof(applyProtection)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(alignment))
                    ctObj.alignment = CT_CellAlignment.Parse(childNode, namespaceManager);
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
            XmlHelper.WriteAttribute(sw, nameof(numFmtId), this.numFmtId, true);
            XmlHelper.WriteAttribute(sw, nameof(fontId), this.fontId, true);
            XmlHelper.WriteAttribute(sw, nameof(fillId), this.fillId, true);
            XmlHelper.WriteAttribute(sw, nameof(borderId), this.borderId, true);
            XmlHelper.WriteAttribute(sw, nameof(xfId), this.xfId, true);
            XmlHelper.WriteAttribute(sw, nameof(quotePrefix), this.quotePrefix, false);
            XmlHelper.WriteAttribute(sw, nameof(pivotButton), this.pivotButton, false);
            if (this.applyNumberFormat)
                XmlHelper.WriteAttribute(sw, nameof(applyNumberFormat), this.applyNumberFormat);
            XmlHelper.WriteAttribute(sw, nameof(applyFont), this.applyFont, false);
            if (this.applyBorder)
                XmlHelper.WriteAttribute(sw, nameof(applyBorder), this.applyBorder, true);
            if (this.applyFill)
                XmlHelper.WriteAttribute(sw, nameof(applyFill), this.applyFill);
            XmlHelper.WriteAttribute(sw, nameof(applyAlignment), this.applyAlignment, true);
            if (this.applyProtection)
                XmlHelper.WriteAttribute(sw, nameof(applyProtection), this.applyProtection, true);
            if (this.alignment == null && this.protection == null && this.extLst == null)
            {
                sw.Write("/>");
            }
            else
            {
                sw.Write(">");
                this.alignment?.Write(sw, nameof(alignment));
                this.protection?.Write(sw, nameof(protection));
                this.extLst?.Write(sw, nameof(extLst));
                sw.Write(string.Format("</{0}>", nodeName));
            }
        }

        public override string ToString()
        {
            XmlSerializer serializer = new XmlSerializer(typeof(CT_Xf));
            using (StringWriter stream = new StringWriter())
            {
                serializer.Serialize(stream, this);
                return stream.ToString();
            }
        }

        public bool IsSetFontId()
        {
            return this.fontIdSpecified;
        }
        public bool IsSetAlignment()
        {
            return alignment != null;
        }
        public void UnsetAlignment()
        {
            this.alignment = null;
        }

        public bool IsSetExtLst()
        {
            return this.extLst == null;
        }
        public void UnsetExtLst()
        {
            this.extLst = null;
        }

        public bool IsSetProtection()
        {
            return this.protection != null;
        }
        public void UnsetProtection()
        {
            this.protection = null;
        }
        public bool IsSetLocked()
        {
            // first guess:
            return IsSetProtection() && (protection.locked == true);
        }
        public CT_CellProtection AddNewProtection()
        {
            this.protection = new CT_CellProtection();
            return this.protection;
        }

        [XmlElement]
        public CT_CellAlignment alignment { get; set; } = null;


        [XmlElement]
        public CT_CellProtection protection { get; set; } = null;

        [XmlElement]
        public CT_ExtensionList extLst { get; set; } = null;
        [XmlAttribute]
        public uint numFmtId
        {
            get { return this.numFmtIdField; }
            set
            {
                this.numFmtIdField = value;
                this.numFmtIdSpecified = true;
            }
        }

        [XmlIgnore]
        public bool numFmtIdSpecified { get; set; } = false;

        [XmlAttribute]
        public uint fontId
        {
            get { return this.fontIdField; }
            set
            {
                this.fontIdField = value;
                this.fontIdSpecified = true;
            }
        }

        [XmlIgnore]
        public bool fontIdSpecified { get; set; } = false;

        [XmlAttribute]
        public uint fillId
        {
            get { return this.fillIdField; }
            set
            {
                this.fillIdField = value;
                this.fillIdSpecified = true;
            }
        }

        [XmlIgnore]
        public bool fillIdSpecified { get; set; } = false;

        [XmlAttribute]
        public uint borderId
        {
            get { return this.borderIdField; }
            set
            {
                this.borderIdField = value;
                borderIdSpecified = true;
            }
        }
        [XmlIgnore]
        public bool borderIdSpecified { get; set; } = false;

        [XmlAttribute]
        public uint xfId
        {
            get { return this.xfIdField; }
            set
            {
                this.xfIdField = value;
                this.xfIdSpecified = true;
            }
        }
        [XmlIgnore]
        public bool xfIdSpecified { get; set; } = false;

        [XmlAttribute]
        [DefaultValue(false)]
        public bool quotePrefix { get; set; } = false;

        [XmlAttribute]
        [DefaultValue(false)]
        public bool pivotButton { get; set; } = false;

        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyNumberFormat { get; set; } = false;

        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyFont { get; set; } = false;

        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyFill { get; set; } = false;
        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyBorder { get; set; } = false;
        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyAlignment { get; set; } = false;

        [XmlAttribute]
        [DefaultValue(false)]
        public bool applyProtection { get; set; } = false;

        public bool IsSetApplyFill()
        {
            return this.applyFill;
        }
    }
}
