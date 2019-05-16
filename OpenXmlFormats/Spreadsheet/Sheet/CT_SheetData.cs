using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.Xml;
using System.IO;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_SheetData
    {

        public static CT_SheetData Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            return new CT_SheetData
            {
                row = node.ChildNodes.Cast<XmlNode>()
                    .Where(childNode => childNode.LocalName == nameof(row))
                    .Select(childNode => CT_Row.Parse(childNode, namespaceManager))
                    .ToList()
            };

        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            sw.Write(">");
            this.row?.ForEach(x => x.Write(sw, nameof(row)));
            sw.Write($"</{nodeName}>");
        }

        //public CT_SheetData()
        //{
        //    this.rowField = new List<CT_Row>();
        //}
        public CT_Row AddNewRow()
        {
            this.row = this.row ?? new List<CT_Row>();
            var newRow = new CT_Row();
            row.Add(newRow);
            return newRow;
        }
        public CT_Row InsertNewRow(int index)
        {
            this.row = this.row ?? new List<CT_Row>();
            var newRow = new CT_Row();
            row.Insert(index, newRow);
            return newRow;
        }
        public void RemoveRows(IList<CT_Row> toRemove)
        {
            if (row == null) return;
            foreach (CT_Row r in toRemove)
            {
                row.Remove(r);
            }
        }
        public void RemoveRow(int rowNum)
        {
            var rowToRemove = this.row?.FirstOrDefault(ctrow => ctrow.r == rowNum);
            row.Remove(rowToRemove);
        }
        public int SizeOfRowArray()
        {
            return row?.Count ?? 0;
        }

        public CT_Row GetRowArray(int index)
        {
            return row?[index];
        }
        [XmlElement(nameof(row))]
        public List<CT_Row> row { get; set; } = null;
        [XmlIgnore]
        public bool rowSpecified
        {
            get { return null != row; }
        }
    }
}
