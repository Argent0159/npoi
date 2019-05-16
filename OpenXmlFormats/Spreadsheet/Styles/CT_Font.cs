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
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ElementName = "font")]
    public class CT_Font
    {
        public CT_Font()
        {
        }

        public static CT_Font Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Font ctObj = new CT_Font();
            ctObj.charset = new List<CT_IntProperty>();
            ctObj.family = new List<CT_IntProperty>();
            ctObj.b = new List<CT_BooleanProperty>();
            ctObj.i = new List<CT_BooleanProperty>();
            ctObj.strike = new List<CT_BooleanProperty>();
            ctObj.color = new List<CT_Color>();
            ctObj.sz = new List<CT_FontSize>();
            ctObj.u = new List<CT_UnderlineProperty>();
            ctObj.vertAlign = new List<CT_VerticalAlignFontProperty>();
            ctObj.scheme = new List<CT_FontScheme>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "outline")
                    ctObj.outline = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "shadow")
                    ctObj.shadow = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "condense")
                    ctObj.condense = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "extend")
                    ctObj.extend = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "name")
                    ctObj.name= CT_FontName.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == "charset")
                    ctObj.charset.Add(CT_IntProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "family")
                    ctObj.family.Add(CT_IntProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "b")
                    ctObj.b.Add(CT_BooleanProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "i")
                    ctObj.i.Add(CT_BooleanProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "strike")
                    ctObj.strike.Add(CT_BooleanProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "color")
                    ctObj.color.Add(CT_Color.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "sz")
                    ctObj.sz.Add(CT_FontSize.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "u")
                    ctObj.u.Add(CT_UnderlineProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "vertAlign")
                    ctObj.vertAlign.Add(CT_VerticalAlignFontProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == "scheme")
                    ctObj.scheme.Add(CT_FontScheme.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            if (this.b != null)
            {
                foreach (CT_BooleanProperty x in this.b)
                {
                    x.Write(sw, "b");
                }
            }
            if (this.i != null)
            {
                foreach (CT_BooleanProperty x in this.i)
                {
                    x.Write(sw, "i");
                }
            }

            if (this.strike != null)
            {
                foreach (CT_BooleanProperty x in this.strike)
                {
                    x.Write(sw, "strike");
                }
            }
            if (this.condense != null)
                this.condense.Write(sw, "condense");
            if (this.extend != null)
                this.extend.Write(sw, "extend");
            if (this.outline != null)
                this.outline.Write(sw, "outline");
            if (this.shadow != null)
                this.shadow.Write(sw, "shadow");
            if (this.u != null)
            {
                foreach (CT_UnderlineProperty x in this.u)
                {
                    x.Write(sw, "u");
                }
            }
            if (this.vertAlign != null)
            {
                foreach (CT_VerticalAlignFontProperty x in this.vertAlign)
                {
                    x.Write(sw, "vertAlign");
                }
            }

            if (this.sz != null)
            {
                foreach (CT_FontSize x in this.sz)
                {
                    x.Write(sw, "sz");
                }
            }

            if (this.color != null)
            {
                foreach (CT_Color x in this.color)
                {
                    x.Write(sw, "color");
                }
            }
            if (this.name != null)
                this.name.Write(sw, "name");

            if (this.family != null)
            {
                foreach (CT_IntProperty x in this.family)
                {
                    x.Write(sw, "family");
                }
            }
            if (this.charset != null)
            {
                foreach (CT_IntProperty x in this.charset)
                {
                    x.Write(sw, "charset");
                }
            }
            if (this.scheme != null)
            {
                foreach (CT_FontScheme x in this.scheme)
                {
                    x.Write(sw, "scheme");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }


        //public static string GetString(CT_Font font)
        //{
        //    using (StringWriter writer = new StringWriter())
        //    {
        //        serializer.Serialize(writer, font, namespaces);
        //        return writer.ToString();
        //    }
        //}
        #region name
        [XmlElement]
        public CT_FontName name { get; set; } = null;
        public int sizeOfNameArray()
        {
            if (this.name == null)
                return 0;
            return 1;
        }
        public CT_FontName AddNewName()
        {
            this.name = new CT_FontName();
            return this.name;
        }
        #endregion name

        #region charset
        [XmlElement]
        public List<CT_IntProperty> charset { get; set; } = null;
        public int sizeOfCharsetArray()
        {
            if (this.charset == null)
                return 0;
            return this.charset.Count;
        }
        public CT_IntProperty AddNewCharset()
        {
            if (this.charset == null)
                this.charset = new List<CT_IntProperty>();
            CT_IntProperty prop = new CT_IntProperty();
            this.charset.Add(prop);
            return prop;
        }
        public void SetCharsetArray(int index, CT_IntProperty value)
        {
            this.charset[index] = value;
        }
        public CT_IntProperty GetCharsetArray(int index)
        {
            return this.charset[index];
        }
        #endregion charset

        #region family
        [XmlElement]
        public List<CT_IntProperty> family { get; set; } = null;
        public int sizeOfFamilyArray()
        {
            if (this.family == null)
                return 0;
            return this.family.Count;
        }
        public CT_IntProperty AddNewFamily()
        {
            if (this.family == null)
                this.family = new List<CT_IntProperty>();
            CT_IntProperty newfamily = new CT_IntProperty();
            this.family.Add(newfamily);
            return newfamily;
        }
        public void SetFamilyArray(int index, CT_IntProperty value)
        {
            this.family[index] = value;
        }
        public CT_IntProperty GetFamilyArray(int index)
        {
            return this.family[index];
        }
        #endregion family

        #region b
        [XmlElement]
        public List<CT_BooleanProperty> b { get; set; } = null;
        public int SizeOfBArray()
        {
            if (this.b == null)
                return 0;
            return this.b.Count;
        }
        public CT_BooleanProperty AddNewB()
        {
            if (this.b == null)
                this.b = new List<CT_BooleanProperty>();
            CT_BooleanProperty newB = new CT_BooleanProperty();
            this.b.Add(newB);
            return newB;
        }
        public void SetBArray(int index, CT_BooleanProperty value)
        {
            this.b[index] = value;
        }
        public void SetBArray(List<CT_BooleanProperty> array)
        {
            this.b = array;
        }
        public CT_BooleanProperty GetBArray(int index)
        {
            return this.b[index];
        }
        #endregion b

        #region i
        [XmlElement]
        public List<CT_BooleanProperty> i { get; set; } = null;
        public int sizeOfIArray()
        {
            if (this.i == null)
                return 0;
            return this.i.Count;
        }
        public CT_BooleanProperty AddNewI()
        {
            if (this.i == null)
                this.i = new List<CT_BooleanProperty>();
            CT_BooleanProperty newI = new CT_BooleanProperty();
            this.i.Add(newI);
            return newI;
        }
        public void SetIArray(int index, CT_BooleanProperty value)
        {
            this.i[index] = value;
        }
        public void SetIArray(List<CT_BooleanProperty> array)
        {
            this.i = array;
        }
        public CT_BooleanProperty GetIArray(int index)
        {
            return this.i[index];
        }
        #endregion i

        #region strike
        [XmlElement]
        public List<CT_BooleanProperty> strike { get; set; } = null;
        public int sizeOfStrikeArray()
        {
            if (this.strike == null)
                return 0;
            return this.strike.Count;
        }
        public CT_BooleanProperty AddNewStrike()
        {
            if (this.strike == null)
                this.strike = new List<CT_BooleanProperty>();
            CT_BooleanProperty prop = new CT_BooleanProperty();
            this.strike.Add(prop);
            return prop;
        }
        public void SetStrikeArray(int index, CT_BooleanProperty value)
        {
            this.strike[index] = value;
        }
        public void SetStrikeArray(List<CT_BooleanProperty> array)
        {
            this.strike = array;
        }
        public CT_BooleanProperty GetStrikeArray(int index)
        {
            return this.strike[index];
        }
        #endregion strike

        #region outline
        [XmlElement]
        public CT_BooleanProperty outline { get; set; } = null;
        public int sizeOfOutlineArray()
        {
            return this.outline == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewOutline()
        {
            this.outline = new CT_BooleanProperty();
            return this.outline;
        }
        public void SetOutlineArray(CT_BooleanProperty[] array)
        {
            this.outline = array.Length > 0 ? array[0] : null;
        }
        public CT_BooleanProperty GetOutlineArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.outline;
        }
        #endregion outline

        #region shadow
        [XmlElement]
        public CT_BooleanProperty shadow { get; set; } = null;
        public int sizeOfShadowArray()
        {
            if (this.shadow == null)
                return 0;
            return this.shadow == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewShadow()
        {
            this.shadow = new CT_BooleanProperty();
            return this.shadow;
        }
        public CT_BooleanProperty GetShadowArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.shadow;
        }
        #endregion shadow

        #region condense
        [XmlElement]
        public CT_BooleanProperty condense { get; set; } = null;
        public int sizeOfCondenseArray()
        {
            if (this.condense == null)
                return 0;
            return this.condense == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewCondense()
        {
            this.condense = new CT_BooleanProperty();
            return this.condense;
        }
        public CT_BooleanProperty GetCondenseArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.condense;
        }
        #endregion condense

        #region extend
        [XmlElement]
        public CT_BooleanProperty extend { get; set; } = null;
        public int sizeOfExtendArray()
        {
            return this.extend == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewExtend()
        {
            this.extend = new CT_BooleanProperty();
            return this.extend;
        }
        public CT_BooleanProperty GetExtendArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.extend;
        }
        #endregion extend

        #region color
        [XmlElement]
        public List<CT_Color> color { get; set; } = null;
        public int sizeOfColorArray()
        {
            if (this.color == null)
                return 0;
            return this.color.Count;
        }
        public CT_Color GetColorArray(int index)
        {
            return this.color[index];
        }
        public void SetColorArray(int index, CT_Color value)
        {
            this.color[index] = value;
        }
        public void SetColorArray(List<CT_Color> array)
        {
            this.color = array;
        }
        public CT_Color AddNewColor()
        {
            if (this.color == null)
                this.color = new List<CT_Color>();
            CT_Color newColor = new CT_Color();
            this.color.Add(newColor);
            return newColor;
        }
        #endregion color

        #region sz
        [XmlElement]
        public List<CT_FontSize> sz { get; set; } = null;
        public int sizeOfSzArray()
        {
            if (this.sz == null)
                return 0;
            return this.sz.Count;
        }
        public CT_FontSize AddNewSz()
        {
            if (this.sz == null)
                this.sz = new List<CT_FontSize>();
            CT_FontSize newFs = new CT_FontSize();
            this.sz.Add(newFs);
            return newFs;
        }
        public void SetSzArray(int index, CT_FontSize value)
        {
            this.sz[index] = value;
        }
        public void SetSzArray(List<CT_FontSize> array)
        {
            this.sz = array;
        }
        public CT_FontSize GetSzArray(int index)
        {
            return this.sz[index];
        }
        #endregion sz

        #region u
        [XmlElement]
        public List<CT_UnderlineProperty> u { get; set; } = null;
        public int sizeOfUArray()
        {
            if (this.u == null)
                return 0;
            return this.u.Count;
        }
        public CT_UnderlineProperty AddNewU()
        {
            if (this.u == null)
                this.u = new List<CT_UnderlineProperty>();
            CT_UnderlineProperty newU = new CT_UnderlineProperty();
            this.u.Add(newU);
            return newU;
        }
        public void SetUArray(int index, CT_UnderlineProperty value)
        {
            if (u[index] != null)
                this.u[index] = value;
            else
                this.u.Insert(index, value);
        }
        public void SetUArray(List<CT_UnderlineProperty> array)
        {
            this.u = array;
        }
        public CT_UnderlineProperty GetUArray(int index)
        {
            return this.u[index];
        }
        #endregion u

        #region vertAlign
        [XmlElement]
        public List<CT_VerticalAlignFontProperty> vertAlign { get; set; } = null;
        public int sizeOfVertAlignArray()
        {
            if (this.vertAlign == null)
                return 0;
            return this.vertAlign.Count;
        }
        public CT_VerticalAlignFontProperty AddNewVertAlign()
        {
            if (this.vertAlign == null)
                this.vertAlign = new List<CT_VerticalAlignFontProperty>();
            CT_VerticalAlignFontProperty prop = new CT_VerticalAlignFontProperty();
            this.vertAlign.Add(prop);
            return prop;
        }
        public void SetVertAlignArray(int index, CT_VerticalAlignFontProperty value)
        {
            this.vertAlign[index] = value;
        }
        public void SetVertAlignArray(List<CT_VerticalAlignFontProperty> array)
        {
            this.vertAlign =array;
        }
        public CT_VerticalAlignFontProperty GetVertAlignArray(int index)
        {
            return this.vertAlign[index];
        }
        #endregion vertAlign

        #region scheme
        [XmlElement]
        public List<CT_FontScheme> scheme { get; set; } = null;
        public int sizeOfSchemeArray()
        {
            if (this.scheme == null)
                return 0;
            return this.scheme.Count;
        }
        public CT_FontScheme AddNewScheme()
        {
            if (this.scheme == null)
                this.scheme = new List<CT_FontScheme>();
            CT_FontScheme newScheme = new CT_FontScheme();
            this.scheme.Add(newScheme);
            return newScheme;
        }
        public void SetSchemeArray(int index, CT_FontScheme value)
        {
            this.scheme[index] = value;
        }
        public CT_FontScheme GetSchemeArray(int index)
        {
            return this.scheme[index];
        }
        #endregion scheme

        public override string ToString()
        {
            using (MemoryStream ms = new MemoryStream())
            {
                StreamWriter sw = new StreamWriter(ms);
                this.Write(sw, "font");
                sw.Flush();
                ms.Position = 0;
                StreamReader sr = new StreamReader(ms);
                string result = sr.ReadToEnd();
                return result;
            }
        }

        public CT_Font Clone()
        {
            CT_Font ctFont = new CT_Font();

            if (this.name!=null)
            {
                  CT_FontName newName = ctFont.AddNewName();
                    newName.val = this.name.val;
            }
            if (this.charset!=null)
            {
                foreach (CT_IntProperty ctCharset in this.charset)
                {
                    CT_IntProperty newCharset = ctFont.AddNewCharset();
                    newCharset.val = ctCharset.val;
                }
            }
            if (this.family!=null)
            {
                foreach (CT_IntProperty ctFamily in this.family)
                {
                    CT_IntProperty newFamily = ctFont.AddNewFamily();
                    newFamily.val = ctFamily.val;
                }
            }
            if (this.b != null)
            {
                foreach (CT_BooleanProperty ctB in this.b)
                {
                    CT_BooleanProperty newB = ctFont.AddNewB();
                    newB.val = ctB.val;
                }
            }
            if (this.i != null)
            {
                foreach (CT_BooleanProperty ctI in this.i)
                {
                    CT_BooleanProperty newI = ctFont.AddNewI();
                    newI.val = ctI.val;
                }
            }
            if (this.strike != null)
            {
                foreach (CT_BooleanProperty ctStrike in this.strike)
                {
                    CT_BooleanProperty newstrike = ctFont.AddNewStrike();
                    newstrike.val = ctStrike.val;
                }
            }
            if (this.outline != null)
            {
                ctFont.outline = new CT_BooleanProperty();
                ctFont.outline.val = this.outline.val;
            }
            if (this.shadow != null)
            {
                ctFont.shadow = new CT_BooleanProperty();
                ctFont.shadow.val = this.shadow.val;
            }
            if (this.condense != null)
            {
                ctFont.condense = new CT_BooleanProperty();
                ctFont.condense.val = this.condense.val;
            }
            if (this.extend != null)
            {
                ctFont.extend = new CT_BooleanProperty();
                ctFont.extend.val = this.extend.val;
            }
            if (this.color != null)
            {
                foreach (CT_Color ctColor in this.color)
                {
                    CT_Color newColor = ctFont.AddNewColor();
                    newColor.theme = ctColor.theme; //Forces themeSpecified to true even if a theme wasn't specified.
                    newColor.themeSpecified = ctColor.themeSpecified;
                    newColor.rgb = ctColor.rgb;
                    newColor.rgbSpecified = ctColor.rgbSpecified;
                    newColor.tint = ctColor.tint;
                    newColor.tintSpecified = ctColor.tintSpecified;
                    newColor.auto = ctColor.auto;
                    newColor.autoSpecified = ctColor.autoSpecified;
                    //Does not copy indexed color field because we don't support indexed colors for XSSF.
                    //If copying indexed colors between two documents you need to account for the color palettes
                    //potentially being different between two documents. (MSSQL Reporting Services did this in HSSF)
                }
            }
            if (this.sz != null)
            {
                foreach (CT_FontSize ctSz in this.sz)
                {
                    CT_FontSize newSz = ctFont.AddNewSz();
                    newSz.val = ctSz.val;
                }
            }
            if (this.u != null)
            {
                foreach (CT_UnderlineProperty ctU in this.u)
                {
                    CT_UnderlineProperty newU = ctFont.AddNewU();
                    newU.val = ctU.val;
                }
            }
            if (this.vertAlign != null)
            {
                foreach (CT_VerticalAlignFontProperty ctVertAlign in this.vertAlign)
                {
                    CT_VerticalAlignFontProperty newVertAlign = ctFont.AddNewVertAlign();
                    newVertAlign.val = ctVertAlign.val;
                }

            }
            if (this.scheme != null)
            {
                foreach (CT_FontScheme ctScheme in this.scheme)
                {
                    CT_FontScheme newScheme = ctFont.AddNewScheme();
                    newScheme.val = ctScheme.val;
                }
            }
            return ctFont;
        }
    }
}
