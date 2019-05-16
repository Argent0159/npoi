using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;
using NPOI.OpenXml4Net.Util;
using System.IO;
using System.Xml;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    /// <summary>
    /// Rich Text Run container.
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_RElt
    {
        public static CT_RElt Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RElt ctObj = new CT_RElt();
            XmlNode tNode = node.SelectSingleNode("d:t", namespaceManager);
            if(tNode!=null)
                ctObj.t = tNode.InnerText.Replace("\r", ""); ;
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(rPr))
                    ctObj.rPr = CT_RPrElt.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}>");
            this.rPr?.Write(sw, nameof(rPr));
            if (this.t != null)
                sw.Write($"<{nameof(t)} xml:space=\"preserve\">{XmlHelper.ExcelEncodeString(XmlHelper.EncodeXml(this.t))}</{nameof(t)}>");
            sw.Write($"</{nodeName}>");
        }


        public CT_RPrElt AddNewRPr()
        {
            this.rPr = new CT_RPrElt();
            return rPr;
        }

        /// <summary>
        /// Run Properties
        /// </summary>
        [XmlElement(nameof(rPr))]
        public CT_RPrElt rPr { get; set; } = null;

        /// <summary>
        /// Text
        /// </summary>
        [XmlElement(nameof(t))]
        public string t { get; set; } = string.Empty;
    }
}
