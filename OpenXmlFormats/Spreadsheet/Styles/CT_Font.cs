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
            var ctObj = new CT_Font
            {
                charset = new List<CT_IntProperty>(),
                family = new List<CT_IntProperty>(),
                b = new List<CT_BooleanProperty>(),
                i = new List<CT_BooleanProperty>(),
                strike = new List<CT_BooleanProperty>(),
                color = new List<CT_Color>(),
                sz = new List<CT_FontSize>(),
                u = new List<CT_UnderlineProperty>(),
                vertAlign = new List<CT_VerticalAlignFontProperty>(),
                scheme = new List<CT_FontScheme>()
            };
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(outline))
                    ctObj.outline = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(shadow))
                    ctObj.shadow = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(condense))
                    ctObj.condense = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extend))
                    ctObj.extend = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(name))
                    ctObj.name = CT_FontName.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(charset))
                    ctObj.charset.Add(CT_IntProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == nameof(family))
                    ctObj.family.Add(CT_IntProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == nameof(b))
                    ctObj.b.Add(CT_BooleanProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == nameof(i))
                    ctObj.i.Add(CT_BooleanProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == nameof(strike))
                    ctObj.strike.Add(CT_BooleanProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == nameof(color))
                    ctObj.color.Add(CT_Color.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == nameof(sz))
                    ctObj.sz.Add(CT_FontSize.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == nameof(u))
                    ctObj.u.Add(CT_UnderlineProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == nameof(vertAlign))
                    ctObj.vertAlign.Add(CT_VerticalAlignFontProperty.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == nameof(scheme))
                    ctObj.scheme.Add(CT_FontScheme.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            this.b?.ForEach(x => x.Write(sw, nameof(b)));
            this.i?.ForEach(x => x.Write(sw, nameof(i)));
            this.strike?.ForEach(x => x.Write(sw, nameof(strike)));
            this.condense?.Write(sw, nameof(condense));
            this.extend?.Write(sw, nameof(extend));
            this.outline?.Write(sw, nameof(outline));
            this.shadow?.Write(sw, nameof(shadow));
            this.u?.ForEach(x => x.Write(sw, nameof(u)));
            this.vertAlign?.ForEach(x => x.Write(sw, nameof(vertAlign)));
            this.sz?.ForEach(x=>x.Write(sw, nameof(sz)));
            this.color?.ForEach(x => x.Write(sw, nameof(color)));
            this.name?.Write(sw, nameof(name));
            this.family?.ForEach(x => x.Write(sw, nameof(family)));
            this.charset?.ForEach(x => x.Write(sw, nameof(charset)));
            this.scheme?.ForEach(x => x.Write(sw, nameof(scheme)));
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
            return this.name == null ? 0 : 1;
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
            return this.charset?.Count ?? 0;
        }
        public CT_IntProperty AddNewCharset()
        {
            this.charset = this.charset ?? new List<CT_IntProperty>();
            var prop = new CT_IntProperty();
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
            return this.family?.Count ?? 0;
        }
        public CT_IntProperty AddNewFamily()
        {
            this.family = this.family ?? new List<CT_IntProperty>();
            var newfamily = new CT_IntProperty();
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
            return this.b?.Count ?? 0;
        }
        public CT_BooleanProperty AddNewB()
        {
            this.b = this.b ?? new List<CT_BooleanProperty>();
            var newB = new CT_BooleanProperty();
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
            return this.i?.Count ?? 0;
        }
        public CT_BooleanProperty AddNewI()
        {
            this.i = this.i ?? new List<CT_BooleanProperty>();
            var newI = new CT_BooleanProperty();
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
            return this.strike?.Count ?? 0;
        }
        public CT_BooleanProperty AddNewStrike()
        {
            this.strike = this.strike ?? new List<CT_BooleanProperty>();
            var prop = new CT_BooleanProperty();
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
            return this.color?.Count ?? 0;
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
            this.color = this.color ?? new List<CT_Color>();
            var newColor = new CT_Color();
            this.color.Add(newColor);
            return newColor;
        }
        #endregion color

        #region sz
        [XmlElement]
        public List<CT_FontSize> sz { get; set; } = null;
        public int sizeOfSzArray()
        {
            return this.sz?.Count ?? 0;
        }
        public CT_FontSize AddNewSz()
        {
            this.sz = this.sz ?? new List<CT_FontSize>();
            var newFs = new CT_FontSize();
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
            return this.u?.Count ?? 0;
        }
        public CT_UnderlineProperty AddNewU()
        {
            this.u = this.u ?? new List<CT_UnderlineProperty>();
            var newU = new CT_UnderlineProperty();
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
            return this.vertAlign?.Count ?? 0;
        }
        public CT_VerticalAlignFontProperty AddNewVertAlign()
        {
            this.vertAlign = this.vertAlign ?? new List<CT_VerticalAlignFontProperty>();
            var prop = new CT_VerticalAlignFontProperty();
            this.vertAlign.Add(prop);
            return prop;
        }
        public void SetVertAlignArray(int index, CT_VerticalAlignFontProperty value)
        {
            this.vertAlign[index] = value;
        }
        public void SetVertAlignArray(List<CT_VerticalAlignFontProperty> array)
        {
            this.vertAlign = array;
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
            return this.scheme?.Count ?? 0;
        }
        public CT_FontScheme AddNewScheme()
        {
            this.scheme = this.scheme ?? new List<CT_FontScheme>();
            var newScheme = new CT_FontScheme();
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
            using (var ms = new MemoryStream())
            {
                var sw = new StreamWriter(ms);
                this.Write(sw, "font");
                sw.Flush();
                ms.Position = 0;
                var sr = new StreamReader(ms);
                string result = sr.ReadToEnd();
                return result;
            }
        }

        public CT_Font Clone()
        {
            var ctFont = new CT_Font();

            if (this.name != null)
            {
                CT_FontName newName = ctFont.AddNewName();
                newName.val = this.name.val;
            }
            if (this.charset != null)
            {
                foreach (CT_IntProperty ctCharset in this.charset)
                {
                    CT_IntProperty newCharset = ctFont.AddNewCharset();
                    newCharset.val = ctCharset.val;
                }
            }
            if (this.family != null)
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
