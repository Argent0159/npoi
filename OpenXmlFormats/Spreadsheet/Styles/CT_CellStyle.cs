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
            return new CT_CellStyle
            {
                name = XmlHelper.ReadString(node.Attributes[nameof(name)]),
                xfId = XmlHelper.ReadUInt(node.Attributes[nameof(xfId)]),
                builtinId = XmlHelper.ReadUInt(node.Attributes[nameof(builtinId)]),
                iLevel = XmlHelper.ReadUInt(node.Attributes[nameof(iLevel)]),
                hidden = XmlHelper.ReadBool(node.Attributes[nameof(hidden)]),
                customBuiltin = XmlHelper.ReadBool(node.Attributes[nameof(customBuiltin)]),
                extLst = CT_ExtensionList.Parse(
                    node.ChildNodes.Cast<XmlNode>().Last(
                        childNode => childNode.LocalName == nameof(extLst)), namespaceManager)
            };
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(xfId), this.xfId, true);
            XmlHelper.WriteAttribute(sw, nameof(builtinId), this.builtinId, true);
            XmlHelper.WriteAttribute(sw, nameof(iLevel), this.iLevel);
            XmlHelper.WriteAttribute(sw, nameof(hidden), this.hidden, false);
            XmlHelper.WriteAttribute(sw, nameof(customBuiltin), this.customBuiltin, false);
            if (this.extLst != null)
            {
                sw.Write(">");
                this.extLst.Write(sw, nameof(extLst));
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
