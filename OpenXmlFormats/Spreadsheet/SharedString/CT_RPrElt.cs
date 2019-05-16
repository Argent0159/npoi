using System;
using System.Collections.Generic;
using System.IO;

using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    /// <summary>
    /// Properties of Rich Text Run.
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_RPrElt
    {
        public static CT_RPrElt Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RPrElt ctObj = new CT_RPrElt();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(rFont))
                    ctObj.rFont = CT_FontName.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(charset))
                    ctObj.charset = CT_IntProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(family))
                    ctObj.family = CT_IntProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(b))
                    ctObj.b = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(i))
                    ctObj.i = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(strike))
                    ctObj.strike = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(outline))
                    ctObj.outline = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(shadow))
                    ctObj.shadow = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(condense))
                    ctObj.condense = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extend))
                    ctObj.extend = CT_BooleanProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(color))
                    ctObj.color = CT_Color.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(sz))
                    ctObj.sz = CT_FontSize.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(u))
                    ctObj.u = CT_UnderlineProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(vertAlign))
                    ctObj.vertAlign = CT_VerticalAlignFontProperty.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(scheme))
                    ctObj.scheme = CT_FontScheme.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}>");
            this.sz?.Write(sw, nameof(sz));
            this.color?.Write(sw, nameof(color));
            this.rFont?.Write(sw, nameof(rFont));
            this.family?.Write(sw, nameof(family));
            this.charset?.Write(sw, nameof(charset));
            this.b?.Write(sw, nameof(b));
            this.i?.Write(sw, nameof(i));
            this.strike?.Write(sw, nameof(strike));
            this.outline?.Write(sw, nameof(outline));
            this.shadow?.Write(sw, nameof(shadow));
            this.condense?.Write(sw, nameof(condense));
            this.extend?.Write(sw, nameof(extend));
            this.u?.Write(sw, nameof(u));
            this.vertAlign?.Write(sw, nameof(vertAlign));
            this.scheme?.Write(sw, nameof(scheme));
            sw.Write($"</{nodeName}>");
        }


        #region rFont
        [XmlElement]
        public CT_FontName rFont { get; set; } = null;
        [XmlIgnore]
        // do not remove this field or change the name, because it is automatically used by the XmlSerializer to decide if the name attribute should be printed or not.
        public bool rFontSpecified
        {
            get { return (null != rFont); }
        }
        public int SizeOfRFontArray()
        {
            return this.rFont == null ? 0 : 1;
        }
        public CT_FontName AddNewRFont()
        {
            this.rFont = new CT_FontName();
            return this.rFont;
        }
        public CT_FontName GetRFontArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.rFont;
        }
        #endregion rFont

        #region charset
        [XmlElement]
        public CT_IntProperty charset { get; set; } = null;
        [XmlIgnore]
        public bool charsetSpecified
        {
            get { return (null != charset); }
        }
        public int sizeOfCharsetArray()
        {
            return this.charset == null ? 0 : 1;
        }
        public CT_IntProperty AddNewCharset()
        {
            this.charset = new CT_IntProperty();
            return this.charset;
        }
        public CT_IntProperty GetCharsetArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.charset;
        }
        #endregion charset

        #region family
        [XmlElement]
        public CT_IntProperty family { get; set; } = null;
        [XmlIgnore]
        public bool familySpecified
        {
            get { return (null != family); }
        }
        public int SizeOfFamilyArray()
        {
            return this.family == null ? 0 : 1;
        }
        public CT_IntProperty AddNewFamily()
        {
            this.family = new CT_IntProperty();
            return this.family;
        }
        //public void SetFamilyArray()
        //{
        //    this.familyField = null;
        //}
        public CT_IntProperty GetFamilyArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.family;
        }
        #endregion family

        #region b
        [XmlElement]
        public CT_BooleanProperty b { get; set; } = null;
        [XmlIgnore]
        public bool bSpecified
        {
            get { return (null != b); }
        }
        public int SizeOfBArray()
        {
            return this.b == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewB()
        {
            this.b = new CT_BooleanProperty();
            return this.b;
        }
        public void SetBArray(CT_BooleanProperty[] array)
        {
            this.b = array.Length > 0 ? array[0] : null;
        }
        public CT_BooleanProperty GetBArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.b;
        }
        #endregion b

        #region i
        [XmlElement]
        public CT_BooleanProperty i { get; set; } = null;
        [XmlIgnore]
        public bool iSpecified
        {
            get { return (null != i); }
        }
        public int SizeOfIArray()
        {
            return this.i == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewI()
        {
            this.i = new CT_BooleanProperty();
            return this.i;
        }
        public void SetIArray(CT_BooleanProperty[] array)
        {
            this.i = array.Length > 0 ? array[0] : null;
        }
        public CT_BooleanProperty GetIArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.i;
        }
        #endregion i

        #region strike
        [XmlElement]
        public CT_BooleanProperty strike { get; set; } = null;
        [XmlIgnore]
        public bool strikeSpecified
        {
            get { return (null != strike); }
        }
        public int sizeOfStrikeArray()
        {
            return this.strike == null ? 0 : 1;
        }
        public CT_BooleanProperty AddNewStrike()
        {
            this.strike = new CT_BooleanProperty();
            return this.strike;
        }
        public void SetStrikeArray(CT_BooleanProperty[] array)
        {
            this.strike = array.Length > 0 ? array[0] : null;
        }
        public CT_BooleanProperty GetStrikeArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.strike;
        }
        #endregion strike

        #region outline
        [XmlElement]
        public CT_BooleanProperty outline { get; set; } = null;
        [XmlIgnore]
        public bool outlineSpecified
        {
            get { return (null != outline); }
        }
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
        [XmlIgnore]
        public bool shadowSpecified
        {
            get { return (null != shadow); }
        }
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
        [XmlIgnore]
        public bool condenseSpecified
        {
            get { return (null != condense); }
        }
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
        [XmlIgnore]
        public bool extendSpecified
        {
            get { return (null != extend); }
        }
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
        public CT_Color color { get; set; } = null;
        [XmlIgnore]
        public bool colorSpecified
        {
            get { return (null != color); }
        }
        public int SizeOfColorArray()
        {
            return this.color == null ? 0 : 1;
        }
        public CT_Color GetColorArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.color;
        }
        public void SetColorArray(CT_Color[] array)
        {
            this.color = array.Length > 0 ? array[0] : null;
        }
        public CT_Color AddNewColor()
        {
            this.color = new CT_Color();
            return this.color;
        }
        #endregion color

        #region sz
        [XmlElement]
        public CT_FontSize sz { get; set; } = null;
        [XmlIgnore]
        public bool szSpecified
        {
            get { return (null != sz); }
        }
        public int SizeOfSzArray()
        {
            return this.sz == null ? 0 : 1;
        }
        public CT_FontSize AddNewSz()
        {
            this.sz = new CT_FontSize();
            return this.sz;
        }
        public void SetSzArray(CT_FontSize[] array)
        {
            this.sz = array.Length > 0 ? array[0] : null;
        }
        public CT_FontSize GetSzArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.sz;
        }
        #endregion sz

        #region u
        [XmlElement]
        public CT_UnderlineProperty u { get; set; } = null;
        [XmlIgnore]
        public bool uSpecified
        {
            get { return (null != u); }
        }
        public int SizeOfUArray()
        {
            return this.u == null ? 0 : 1;
        }
        public CT_UnderlineProperty AddNewU()
        {
            this.u = new CT_UnderlineProperty();
            return this.u;
        }
        public void SetUArray(CT_UnderlineProperty[] array)
        {
            this.u = array.Length > 0 ? array[0] : null;
        }
        public CT_UnderlineProperty GetUArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.u;
        }
        #endregion u

        #region vertAlign
        [XmlElement]
        public CT_VerticalAlignFontProperty vertAlign { get; set; } = null;
        [XmlIgnore]
        public bool vertAlignSpecified
        {
            get { return (null != vertAlign); }
        }
        public int sizeOfVertAlignArray()
        {
            return this.vertAlign == null ? 0 : 1;
        }
        public CT_VerticalAlignFontProperty AddNewVertAlign()
        {
            this.vertAlign = new CT_VerticalAlignFontProperty();
            return this.vertAlign;
        }
        public void SetVertAlignArray(CT_VerticalAlignFontProperty[] array)
        {
            this.vertAlign = array.Length > 0 ? array[0] : null;
        }
        public CT_VerticalAlignFontProperty GetVertAlignArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.vertAlign;
        }
        #endregion vertAlign

        #region scheme
        [XmlElement]
        public CT_FontScheme scheme { get; set; } = null;
        [XmlIgnore]
        public bool schemeSpecified
        {
            get { return (null != scheme); }
        }
        public int sizeOfSchemeArray()
        {
            return this.scheme == null ? 0 : 1;
        }
        public CT_FontScheme AddNewScheme()
        {
            this.scheme = new CT_FontScheme();
            return this.scheme;
        }
        public CT_FontScheme GetSchemeArray(int index)
        {
            if (0 != index) { throw new IndexOutOfRangeException("Only an index of 0 is supported"); }
            return this.scheme;
        }
        #endregion scheme

        //public int SizeOfUArray()
        //{
        //    throw new NotImplementedException();
        //}

        //public int SizeOfBArray()
        //{
        //    throw new NotImplementedException();
        //}

        //public int SizeOfSzArray()
        //{
        //    throw new NotImplementedException();
        //}

        //public int SizeOfRFontArray()
        //{
        //    throw new NotImplementedException();
        //}

        //public int SizeOfIArray()
        //{
        //    throw new NotImplementedException();
        //}

        //public int SizeOfColorArray()
        //{
        //    throw new NotImplementedException();
        //}
    }
}
