using System;
using System.Linq;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;
using System.Xml;
using System.IO;

namespace NPOI.OpenXmlFormats.Spreadsheet
{

    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Cols
    {

        //public CT_Cols()
        //{
        //    this.colField = new List<CT_Col>();
        //}
        public void SetColArray(List<CT_Col> array)
        {
            col = array;
        }
        public CT_Col AddNewCol()
        {
            var newCol = new CT_Col();
            this.col.Add(newCol);
            return newCol;
        }
        public CT_Col InsertNewCol(int index)
        {
            var newCol = new CT_Col();
            this.col.Insert(index, newCol);
            return newCol;
        }
        public void RemoveCol(int index)
        {
            this.col.RemoveAt(index);
        }

        public int sizeOfColArray()
        {
            return col.Count;
        }
        public CT_Col GetColArray(int index)
        {
            return col[index];
        }


        public List<CT_Col> GetColList()
        {
            return col;
        }
        [XmlElement]
        public List<CT_Col> col { get; set; } = new List<CT_Col>();

        public static CT_Cols Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            var ctObj = new CT_Cols
            {
                col = node.ChildNodes.Cast<XmlNode>()
                    .Where(n => n.LocalName == nameof(col))
                    .Select(n => CT_Col.Parse(n, namespaceManager))
                    .ToList()
            };
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}");
            sw.Write(">");
            this.col?.ForEach(x => x.Write(sw, nameof(col)));
            sw.Write($"</{nodeName}>");
        }



        public void SetColArray(IEnumerable<CT_Col> colArray)
        {
            this.col = new List<CT_Col>(colArray);
        }
    }
}
