using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class WorkbookDocument
    {
        public WorkbookDocument()
        {
            Workbook = new CT_Workbook();
        }
        public static WorkbookDocument Parse(XmlDocument xmlDoc, XmlNamespaceManager NameSpaceManager)
        {
            var workbook = CT_Workbook.Parse(xmlDoc.DocumentElement, NameSpaceManager);
            return new WorkbookDocument(workbook);
        }
        public WorkbookDocument(CT_Workbook workbook)
        {
            this.Workbook = workbook;
        }
        public CT_Workbook Workbook { get; } = null;
        public void Save(Stream stream)
        {
            using (var sw1 = new StreamWriter(stream))
            {
                Workbook.Write(sw1);
            }
        }
    }
}
