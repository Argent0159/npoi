using NPOI.OpenXml4Net.Util;
using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    //[Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(ElementName = "fonts", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public class CT_Fonts
    {
        public CT_Fonts()
        {
            //this.fontField = new List<CT_Font>();
        }
        public static CT_Fonts Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            var ctObj = new CT_Fonts
            {
                count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]),
                font = node.ChildNodes.Cast<XmlNode>()
                    .Where(childNode => childNode.LocalName == nameof(font))
                    .Select(childNode => CT_Font.Parse(childNode, namespaceManager))
                    .ToList()
            };
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            this.font?.ForEach(x => x.Write(sw, nameof(font)));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public void SetFontArray(List<CT_Font> array)
        {
             font = array;
        }
        [XmlElement]
        public List<CT_Font> font { get; set; }
        [XmlAttribute]
        public uint count { get; set; }

        [XmlIgnore]
        public bool countSpecified { get; set; }
        //internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Fonts));
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
        //    new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });
        //public override string ToString()
        //{
        //    StringWriter stringWriter = new StringWriter();
        //    serializer.Serialize(stringWriter, this, namespaces);
        //    return stringWriter.ToString();
        //}
    }
}
