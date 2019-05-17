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
    public class CT_Colors
    {
        public CT_Colors()
        {
            //this.mruColorsField = new List<CT_Color>();
            //this.indexedColorsField = new List<CT_RgbColor>();
        }
        public static CT_Colors Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Colors ctObj = new CT_Colors();
            
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(indexedColors))
                {
                    ctObj.indexedColors = new List<CT_RgbColor>();
                    foreach (XmlNode c2Node in childNode.ChildNodes)
                    {
                        ctObj.indexedColors.Add(CT_RgbColor.Parse(c2Node, namespaceManager));
                    }
                }
                else if (childNode.LocalName == nameof(mruColors))
                {
                    ctObj.mruColors = new List<CT_Color>();
                    foreach (XmlNode c2Node in childNode.ChildNodes)
                    {
                        ctObj.mruColors.Add(CT_Color.Parse(c2Node, namespaceManager));
                    }
                }
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}>", nodeName));
            if (this.indexedColors != null)
            {
                sw.Write($"<{nameof(indexedColors)}>");
                foreach (CT_RgbColor x in this.indexedColors)
                {
                    x.Write(sw, "rgbColor");
                }
                sw.Write($"</{nameof(indexedColors)}>");
            }
            if (this.mruColors != null)
            {
                sw.Write($"<{nameof(mruColors)}>");
                foreach (CT_Color x in this.mruColors)
                {
                    x.Write(sw, "color");
                }
                sw.Write($"</{nameof(mruColors)}>");
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [XmlArray(Order = 0)]
        [XmlArrayItem("rgbColor", IsNullable = false)]
        public List<CT_RgbColor> indexedColors { get; set; }

        [XmlArray(Order = 1)]
        [XmlArrayItem("color", IsNullable = false)]
        public List<CT_Color> mruColors { get; set; }
    }
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_RgbColor
    {
        public static CT_RgbColor Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RgbColor ctObj = new CT_RgbColor();
            ctObj.rgb = XmlHelper.ReadBytes(node.Attributes[nameof(rgb)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(rgb), this.rgb);
            sw.Write("/>");
        }

        [XmlAttribute(DataType = "hexBinary")]
        // Type ST_UnsignedIntHex is base on xsd:hexBinary, Length 4 (octets!?)
        public byte[] rgb { get; set; } = null;
    }


    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot("color", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Color
    {
        // all attributes are optional
        private bool autoField;

        private uint indexedField;

        private byte[] rgbField; // type ST_UnsignedIntHex is xsd:hexBinary restricted to length 4 (octets!? - see http://www.grokdoc.net/index.php/EOOXML_Objections_Clearinghouse)

        private uint themeField; // TODO change all the uses theme to use uint instead of signed integer variants

        private double tintField;


        #region auto
        [XmlAttribute]
        public bool auto
        {
            get
            {
                return (bool)this.autoField;
            }
            set
            {
                this.autoField = value;
                this.autoSpecified = true;
            }
        }

        [XmlIgnore]
        // do not remove this field or change the name, because it is automatically used by the XmlSerializer to decide if the auto attribute should be printed or not.
        public bool autoSpecified { get; set; } = false;
        public bool IsSetAuto()
        {
            return autoSpecified;
        }

        #endregion auto

        #region indexed
        [XmlAttribute]
        public uint indexed
        {
            get
            {
                return (uint)this.indexedField;
            }
            set
            {
                this.indexedField = value;
                this.indexedSpecified = true;
            }
        }

        [XmlIgnore]
        public bool indexedSpecified { get; set; } = false;
        public bool IsSetIndexed()
        {
            return indexedSpecified;
        }
        #endregion indexed

        #region rgb
        [XmlAttribute(DataType = "hexBinary")]
        // Type ST_UnsignedIntHex is base on xsd:hexBinary
        public byte[] rgb
        {
            get
            {
                return this.rgbField;
            }
            set
            {
                this.rgbField = value;
                this.rgbSpecified = true;
            }
        }

        [XmlIgnore]
        public bool rgbSpecified { get; set; } = false;
        public void SetRgb(byte R, byte G, byte B)
        {
            this.rgbField = new byte[4];
            this.rgbField[0] = 0;
            this.rgbField[1] = R;
            this.rgbField[2] = G;
            this.rgbField[3] = B;
            rgbSpecified = true;
        }
        public bool IsSetRgb()
        {
            return rgbSpecified;
        }
        public void SetRgb(byte[] rgb)
        {
            rgbField = new byte[rgb.Length];
            Array.Copy(rgb, this.rgbField, rgb.Length);
            this.rgbSpecified = true;
        }
        public byte[] GetRgb()
        {
            if (rgbField == null) return null;
            byte[] retVal = new byte[rgbField.Length];
            Array.Copy(rgbField, retVal, rgbField.Length);
            return retVal;
        }
        #endregion rgb

        #region theme
        [XmlAttribute]
        public uint theme
        {
            get
            {
                return (uint)this.themeField;
            }
            set
            {
                this.themeField = value;
                this.themeSpecified = true;
            }
        }

        [XmlIgnore]
        public bool themeSpecified { get; set; }
        public bool IsSetTheme()
        {
            return themeSpecified;
        }
        #endregion theme

        #region tint
        [DefaultValue(0.0D)]
        [XmlAttribute]
        public double tint
        {
            get
            {
                return this.tintField;
            }
            set
            {
                this.tintField = value;
                this.tintSpecified = true;
            }
        }

        [XmlIgnore]
        public bool tintSpecified { get; set; } = false;
        public bool IsSetTint()
        {
            return tintSpecified;
        }
        #endregion tint

        //internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Color));
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
        //    new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });
        //public override string ToString()
        //{
        //    using (StringWriter stringWriter = new StringWriter())
        //    {
        //        serializer.Serialize(stringWriter, this, namespaces);
        //        return stringWriter.ToString();
        //    }
        //}

        public static CT_Color Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Color ctObj = new CT_Color();
            ctObj.auto = XmlHelper.ReadBool(node.Attributes[nameof(auto)]);
            ctObj.autoSpecified = node.Attributes[nameof(auto)] != null;
            ctObj.indexed = XmlHelper.ReadUInt(node.Attributes[nameof(indexed)]);
            ctObj.indexedSpecified = node.Attributes[nameof(indexed)] != null;
            ctObj.rgb = XmlHelper.ReadBytes(node.Attributes[nameof(rgb)]);
            ctObj.rgbSpecified = node.Attributes[nameof(rgb)] != null;
            ctObj.theme = XmlHelper.ReadUInt(node.Attributes[nameof(theme)]);
            ctObj.themeSpecified = node.Attributes[nameof(theme)] != null;
            ctObj.tint = XmlHelper.ReadDouble(node.Attributes[nameof(tint)]);
            ctObj.tintSpecified = node.Attributes[nameof(tint)] != null;
            return ctObj;
        }




        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(auto), this.auto,false);
            if (indexedSpecified)
                XmlHelper.WriteAttribute(sw, nameof(indexed), this.indexed, true);
            if(rgbSpecified)
                XmlHelper.WriteAttribute(sw, nameof(rgb), this.rgb);
            if (themeSpecified)
                XmlHelper.WriteAttribute(sw, nameof(theme), this.theme, true);
            if(tintSpecified)
                XmlHelper.WriteAttribute(sw, nameof(tint), this.tint);
            sw.Write("/>");
        }

        public CT_Color Copy()
        {
            var res = new CT_Color();
            res.autoField = this.autoField;

            res.indexedField = this.indexedField;
            res.indexedSpecified = this.indexedSpecified;

            res.rgbField = this.rgbField == null ? null : (byte[])this.rgbField.Clone(); // type ST_UnsignedIntHex is xsd:hexBinary restricted to length 4 (octets!? - see http://www.grokdoc.net/index.php/EOOXML_Objections_Clearinghouse)
            res.rgbSpecified = this.rgbSpecified;

            res.themeField = this.themeField; // TODO change all the uses theme to use uint instead of signed integer variants
            res.themeSpecified = this.themeSpecified;

            res.tintField = this.tintField;
            res.tintSpecified = this.tintSpecified;

            return res;
        }
    }

}
