using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;

using System.Text;
using System.Xml;
using System.Xml.Serialization;
using NPOI.OpenXml4Net.Util;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class CT_Rst
    {
        public void Set(CT_Rst o)
        {
            this.t = o.t;
            this.r = o.r;
            this.rPh = o.rPh;
            this.phoneticPr = o.phoneticPr;
        }
      



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write($"<{nodeName}>");
            if (this.t != null)
            {
                //TODO: diff has-space case and no-space case
                sw.Write($"<{nameof(t)} xml:space=\"preserve\">{XmlHelper.ExcelEncodeString(XmlHelper.EncodeXml(this.t))}</{nameof(t)}>");
            }

            this.r?.ForEach(x => x.Write(sw, nameof(r)));
            this.rPh?.ForEach(x => x.Write(sw, nameof(rPh)));
            this.phoneticPr?.Write(sw, nameof(phoneticPr));

            sw.Write($"</{nodeName}>");
        }


        #region t
        public bool IsSetT()
        {
            return this.t != null;
        }
        public void unsetT()
        {
            this.t = null;
        }
        [XmlElement(nameof(t), DataType = "string")]
        public string t { get; set; } = null;
        #endregion t

        #region r
        /// <summary>
        /// Rich Text Run
        /// </summary>
        [XmlElement(nameof(r))]
        public List<CT_RElt> r { get; set; } = null;
        private string xmltext;
        [XmlIgnore]
        public string XmlText
        {
            get {
                StringBuilder sb = new StringBuilder();
                using (StringWriter sw = new StringWriter(sb))
                {
                    if (r != null && r.Count > 0)
                    {

                        foreach (CT_RElt r in r)
                        {
                            sw.Write($"<{nameof(r)}>");
                            if (r.rPr != null)
                            {
                                sw.Write($"<{nameof(CT_RElt.rPr)}>");
                                if (r.rPr.b != null && r.rPr.b.val)
                                {
                                    sw.Write($"<{nameof(CT_RPrElt.b)}/>");
                                }
                                if (r.rPr.i != null && r.rPr.i.val)
                                {
                                    sw.Write($"<{nameof(CT_RPrElt.i)}/>");
                                }
                                if (r.rPr.u != null)
                                {
                                    sw.Write($"<{nameof(CT_RPrElt.u)} val=\"{r.rPr.u.val}\"/>");
                                }
                                if (r.rPr.color != null && r.rPr.color.theme > 0)
                                {
                                    sw.Write($"<{nameof(CT_RPrElt.color)} theme=\"{r.rPr.color.theme}\"/>");
                                }
                                if (r.rPr.rFont != null)
                                {
                                    sw.Write($"<{nameof(CT_RPrElt.rFont)} val=\"{r.rPr.rFont.val}\"/>");
                                }
                                if (r.rPr.family != null)
                                {
                                    sw.Write($"<{nameof(CT_RPrElt.family)} val=\"{r.rPr.family.val}\"/>");
                                }
                                if (r.rPr.charset != null)
                                {
                                    sw.Write($"<{nameof(CT_RPrElt.charset)} val=\"{r.rPr.charset.val}\"/>");
                                }
                                if (r.rPr.scheme != null)
                                {
                                    sw.Write($"<{nameof(CT_RPrElt.scheme)} val=\"{r.rPr.scheme.val}\"/>");
                                }
                                if (r.rPr.sz != null)
                                {
                                    sw.Write($"<{nameof(CT_RPrElt.sz)} val=\"{r.rPr.sz.val}\"/>");
                                }
                                if (r.rPr.vertAlign != null)
                                {
                                    sw.Write($"<{nameof(CT_RPrElt.vertAlign)} val=\"{r.rPr.vertAlign.val}\"/>");
                                }
                                sw.Write($"</{nameof(CT_RElt.rPr)}>");
                            }
                            if (r.t != null)
                            {
                                sw.Write($"<{nameof(CT_RElt.t)}");
                                if(r.t.IndexOf(' ')>=0)
                                    sw.Write(" xml:space=\"preserve\"");
                                sw.Write(">");
                                sw.Write(XmlHelper.EncodeXml(r.t));
                                sw.Write($"</{nameof(CT_RElt.t)}>");
                            }
                            sw.Write($"</{nameof(r)}>");
                        }
                    }

                    if (this.t != null)
                    {
                        sw.Write($"<{nameof(t)}");
                        if (this.t.IndexOf(' ') >= 0)
                            sw.Write(" xml:space=\"preserve\"");
                        sw.Write(">");
                        sw.Write(XmlHelper.EncodeXml(this.t));
                        sw.Write($"</{nameof(t)}>");
                    }
                    xmltext = sb.ToString();
                }
                return xmltext; 
            }
            set { xmltext = value; }
        }
        public CT_RElt AddNewR()
        {
            this.r = this.r ?? new List<CT_RElt>();
            CT_RElt r = new CT_RElt();
            this.r.Add(r);
            return r;
        }
        public int sizeOfRArray()
        {
            return r?.Count ?? 0;
        }
        public CT_RElt GetRArray(int index)
        {
            return r?[index];
        }
        #endregion r

        /// <summary>
        /// Phonetic Run
        /// </summary>
        [XmlElement(nameof(rPh))]
        public List<CT_PhoneticRun> rPh { get; set; } = null;
        /// <summary>
        /// Phonetic Properties
        /// </summary>
        [XmlElement(nameof(phoneticPr))]
        public CT_PhoneticPr phoneticPr { get; set; } = null;


        public static CT_Rst Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            CT_Rst ctObj = new CT_Rst();
            ctObj.r = new List<CT_RElt>();
            ctObj.rPh = new List<CT_PhoneticRun>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(phoneticPr))
                    ctObj.phoneticPr = CT_PhoneticPr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(r))
                    ctObj.r.Add(CT_RElt.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == nameof(rPh))
                    ctObj.rPh.Add(CT_PhoneticRun.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == nameof(t))
                    ctObj.t = childNode.InnerText.Replace("\r", "");
            }
            return ctObj;
        }

        public int SizeOfRArray()
        {
            return r.Count;
        }
    }
}
