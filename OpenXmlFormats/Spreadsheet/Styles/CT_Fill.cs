using System;
using System.Collections.Generic;
using System.IO;

using System.Text;
using System.Xml;
using System.Xml.Serialization;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXml4Net.OPC.Internal;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Fill
    {
        [XmlElement]
        public CT_PatternFill patternFill { get; set; } = null;
        [XmlElement]
        public CT_GradientFill gradientFill { get; set; } = null;

        public CT_PatternFill GetPatternFill()
        {
            return this.patternFill;
        }

        public CT_PatternFill AddNewPatternFill()
        {
            this.patternFill = new CT_PatternFill();
            return GetPatternFill();
        }
        public bool IsSetPatternFill()
        {
            return this.patternFill != null;
        }
        public CT_Fill Copy()
        {
            //CT_Fill obj = new CT_Fill();
            //obj.patternFillField = this.patternFillField.Copy();
            //obj.gradientFillField = this.gradientFillField.Copy();
            //return obj;
            return Parse(this.ToString());
        }
        public static CT_Fill Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            var ctObj = new CT_Fill();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(patternFill))
                    ctObj.patternFill = CT_PatternFill.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(gradientFill))
                    ctObj.gradientFill = CT_GradientFill.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            this.patternFill?.Write(sw, nameof(patternFill));
            this.gradientFill?.Write(sw, nameof(gradientFill));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        //internal static XmlSerializer serializer = new XmlSerializer(typeof(CT_Fill));
        //internal static XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces(new XmlQualifiedName[] {
        //    new XmlQualifiedName("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main") });

        public override string ToString()
        {
            using (var ms = new MemoryStream())
            {
                var sw = new StreamWriter(ms);
                this.Write(sw, "fill");
                sw.Flush();
                ms.Position = 0;
                var sr = new StreamReader(ms);
                string result = sr.ReadToEnd();
                return result;
            }
        }

        public static CT_Fill Parse(string p)
        {
            var doc = new XmlDocument();
            doc.LoadXml(p);
            return Parse(doc.DocumentElement, CreateDefaultNSM());
        }
        //TODO: NPOI  duplication code, fix it.
        internal static XmlNamespaceManager CreateDefaultNSM()
        {
            //  Create a NamespaceManager to handle the default namespace, 
            //  and create a prefix for the default namespace:
            var nt = new NameTable();
            var ns = new XmlNamespaceManager(nt);
            ns.AddNamespace(string.Empty, PackageNamespaces.SCHEMA_MAIN);
            ns.AddNamespace("d", PackageNamespaces.SCHEMA_MAIN);
            ns.AddNamespace("a", PackageNamespaces.SCHEMA_DRAWING);
            ns.AddNamespace("xdr", PackageNamespaces.SCHEMA_SHEETDRAWINGS);
            ns.AddNamespace("r", PackageNamespaces.SCHEMA_RELATIONSHIPS);
            ns.AddNamespace("c", PackageNamespaces.SCHEMA_CHART);
            ns.AddNamespace("vt", PackageNamespaces.SCHEMA_VT);
            ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            ns.AddNamespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            ns.AddNamespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            ns.AddNamespace("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            ns.AddNamespace("o", "urn:schemas-microsoft-com:office:office");
            ns.AddNamespace("v", "urn:schemas-microsoft-com:vml");
            ns.AddNamespace("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            // extended properties (app.xml)
            ns.AddNamespace("xp", PackageRelationshipTypes.EXTENDED_PROPERTIES);
            // custom properties
            ns.AddNamespace("ctp", PackageRelationshipTypes.CUSTOM_PROPERTIES);
            // core properties
            ns.AddNamespace("cp", PackagePropertiesPart.NAMESPACE_CP_URI);
            // core property namespaces 
            ns.AddNamespace("dc", PackagePropertiesPart.NAMESPACE_DC_URI);
            ns.AddNamespace("dcterms", PackagePropertiesPart.NAMESPACE_DCTERMS_URI);
            ns.AddNamespace("dcmitype", PackageNamespaces.DCMITYPE);
            ns.AddNamespace("xsi", PackagePropertiesPart.NAMESPACE_XSI_URI);

            ns.AddNamespace("xsd", "http://www.w3.org/2001/XMLSchema");
            return ns;
        }
    }

}
