﻿using NPOI.OpenXml4Net.Util;
using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_CellStyleXfs
    {
        public static CT_CellStyleXfs Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            return new CT_CellStyleXfs
            {
                count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]),
                xf = node.ChildNodes.Cast<XmlNode>()
                    .Where(childNode => childNode.LocalName == nameof(xf))
                    .Select(childNode => CT_Xf.Parse(childNode, namespaceManager))
                    .ToList()
            };
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            this.xf?.ForEach(x => x.Write(sw, nameof(xf)));
            sw.Write($"</{nodeName}>");
        }

        public CT_CellStyleXfs()
        {
            //this.xfField = new List<CT_Xf>();
        }
        public CT_Xf AddNewXf()
        {
            this.xf = this.xf ?? new List<CT_Xf>();
            var xf = new CT_Xf();
            this.xf.Add(xf);
            return xf;
        }
        [XmlElement]
        public List<CT_Xf> xf { get; set; } = null;
        [XmlAttribute]
        public uint count { get; set; }

        [XmlIgnore]
        public bool countSpecified { get; set; }
    }
}
