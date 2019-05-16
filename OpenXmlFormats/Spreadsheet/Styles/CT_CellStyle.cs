﻿using NPOI.OpenXml4Net.Util;
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
    public class CT_CellStyle
    {
        public CT_CellStyle()
        {
           // this.extLstField = new CT_ExtensionList();
        }
        public static CT_CellStyle Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CellStyle ctObj = new CT_CellStyle();
            ctObj.name = XmlHelper.ReadString(node.Attributes["name"]);
            ctObj.xfId = XmlHelper.ReadUInt(node.Attributes["xfId"]);
            ctObj.builtinId = XmlHelper.ReadUInt(node.Attributes["builtinId"]);
            ctObj.iLevel = XmlHelper.ReadUInt(node.Attributes["iLevel"]);
            ctObj.hidden = XmlHelper.ReadBool(node.Attributes["hidden"]);
            ctObj.customBuiltin = XmlHelper.ReadBool(node.Attributes["customBuiltin"]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "extLst")
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, "name", this.name);
            XmlHelper.WriteAttribute(sw, "xfId", this.xfId, true);
            XmlHelper.WriteAttribute(sw, "builtinId", this.builtinId,true);
            XmlHelper.WriteAttribute(sw, "iLevel", this.iLevel);
            XmlHelper.WriteAttribute(sw, "hidden", this.hidden, false);
            XmlHelper.WriteAttribute(sw, "customBuiltin", this.customBuiltin, false);
            if (this.extLst != null)
            {
                sw.Write(">");
                this.extLst.Write(sw, "extLst");
                sw.Write(string.Format("</{0}>", nodeName));
            }
            else
            {
                sw.Write("/>");
            }
        }

        [XmlElement]
        public CT_ExtensionList extLst { get; set; }
        [XmlAttribute]
        public string name { get; set; }
        [XmlAttribute]
        public uint xfId { get; set; }
        [XmlAttribute]
        public uint builtinId { get; set; }

        [XmlIgnore]
        public bool builtinIdSpecified { get; set; }
        [XmlAttribute]
        public uint iLevel { get; set; }

        [XmlIgnore]
        public bool iLevelSpecified { get; set; }
        [XmlAttribute]
        public bool hidden { get; set; }

        [XmlIgnore]
        public bool hiddenSpecified { get; set; }
        [XmlAttribute]
        public bool customBuiltin { get; set; }

        [XmlIgnore]
        public bool customBuiltinSpecified { get; set; }
    }


}
