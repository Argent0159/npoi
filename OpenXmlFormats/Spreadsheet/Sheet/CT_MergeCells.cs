using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.Xml;
using NPOI.OpenXml4Net.Util;
using System.IO;


namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_MergeCells
    {
        public static CT_MergeCells Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;

            return new CT_MergeCells
            {
                count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]),
                mergeCell = node.ChildNodes.Cast<XmlNode>()
                    .Where(childNode => childNode.LocalName == nameof(mergeCell))
                    .Select(childNode => CT_MergeCell.Parse(childNode, namespaceManager))
                    .ToList()
            };
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            mergeCell?.ForEach(x => x?.Write(sw, nameof(mergeCell)));
            sw.Write($"</{nodeName}>");
        }


        public CT_MergeCells()
        {
            this.mergeCell = new List<CT_MergeCell>();
        }
        public CT_MergeCell GetMergeCellArray(int index)
        {
            return this.mergeCell[index];
        }
        public void SetMergeCellArray(IEnumerable<CT_MergeCell> array)
        {
            mergeCell = array.ToList();
        }
        public int sizeOfMergeCellArray()
        {
            return mergeCell.Count;
        }
        public CT_MergeCell AddNewMergeCell()
        {
            var mergecell = new CT_MergeCell();
            mergeCell.Add(mergecell);
            return mergecell;
        }
        [XmlElement]
        public List<CT_MergeCell> mergeCell { get; set; }
        [XmlAttribute]
        public uint count { get; set; }

        [XmlIgnore]
        public bool countSpecified { get; set; }
    }
}
