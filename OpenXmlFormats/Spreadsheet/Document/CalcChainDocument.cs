using NPOI.OpenXml4Net.Util;
using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class CalcChainDocument
    {
        CT_CalcChain calcChain;

        public CalcChainDocument()
        {
            this.calcChain = new CT_CalcChain();
        }
        internal CalcChainDocument(CT_CalcChain calcChain)
        {
            this.calcChain = calcChain;
        }

        public CT_CalcChain GetCalcChain()
        {
            return calcChain;
        }

        public void SetCalcChain(CT_CalcChain calcchain)
        {
            this.calcChain = calcchain;
        }

        public static CalcChainDocument Parse(XmlDocument xmlDoc, XmlNamespaceManager NameSpaceManager)
        {
            var calcChainDoc = new CalcChainDocument();
            foreach (XmlElement node in xmlDoc.SelectNodes("//d:c", NameSpaceManager))
            {
                var cc = new CT_CalcCell();
                if (node.GetAttributeNode(nameof(CT_CalcCell.i))!= null)
                {
                    cc.i = XmlHelper.ReadInt(node.GetAttributeNode(nameof(CT_CalcCell.i)));
                    cc.iSpecified = true;
                }
                cc.r = node.GetAttribute(nameof(CT_CalcCell.r));
                cc.t = XmlHelper.ReadBool(node.GetAttributeNode(nameof(CT_CalcCell.t)));
                cc.s = XmlHelper.ReadBool(node.GetAttributeNode(nameof(CT_CalcCell.s)));
                cc.l = XmlHelper.ReadBool(node.GetAttributeNode(nameof(CT_CalcCell.l)));
                calcChainDoc.calcChain.AddC(cc);
            }
            return calcChainDoc;
        }

        public void Save(Stream stream)
        {
            var sw = new StreamWriter(stream);
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>");
            sw.Write("<calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            foreach (CT_CalcCell cc in calcChain.c)
            {
                sw.Write("<c");
                sw.Write($@" {nameof(CT_CalcCell.r)}=""{cc.r}""");
                if(cc.i>0)
                    sw.Write($@" {nameof(CT_CalcCell.i)}=""{cc.i}""");
                if(cc.s)
                    sw.Write($@" {nameof(CT_CalcCell.s)}=""{Convert.ToInt32(cc.s)}""");
                if (cc.t)
                    sw.Write($@" {nameof(CT_CalcCell.t)}=""{Convert.ToInt32(cc.t)}""");
                if (cc.l)
                    sw.Write($@" {nameof(CT_CalcCell.l)}=""{Convert.ToInt32(cc.l)}""");
                sw.Write("/>");

            }
            sw.Write("</calcChain>");
            sw.Flush();
        }

    }
}
