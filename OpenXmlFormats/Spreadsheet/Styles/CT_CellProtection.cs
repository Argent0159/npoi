using NPOI.OpenXml4Net.Util;
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
    public class CT_CellProtection
    {
        public bool IsSetHidden()
        {
            return this.hidden != false;
        }
        public bool IsSetLocked()
        {
            return this.locked != false;
        }
        public static CT_CellProtection Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            return new CT_CellProtection
            {
                locked = XmlHelper.ReadBool(node.Attributes[nameof(locked)]),
                hidden = XmlHelper.ReadBool(node.Attributes[nameof(hidden)])
            };
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(locked), this.locked);
            if(this.hidden)
                XmlHelper.WriteAttribute(sw, nameof(hidden), this.hidden);
            sw.Write("/>");
        }

        [XmlAttribute]
        public bool locked { get; set; } = false;
        [XmlAttribute]
        public bool hidden { get; set; } = false;
    }
}
