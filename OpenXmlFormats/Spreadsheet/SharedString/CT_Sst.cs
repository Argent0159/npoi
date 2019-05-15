using System;
using System.Collections.Generic;

using System.Text;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [Serializable]
    [System.Diagnostics.DebuggerStepThrough]
    [System.ComponentModel.DesignerCategory("code")]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(ElementName = "sst", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public class CT_Sst
    {
        private int countField;
        private int uniqueCountField;

        public CT_Sst()
        {
            //this.extLstField = new CT_ExtensionList();
            this.si = new List<CT_Rst>();
        }

        [XmlElement]
        public List<CT_Rst> si { get; set; }
        [XmlElement]
        public CT_ExtensionList extLst { get; set; }
        [XmlAttribute]
        public int count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
                this.countSpecified = true;
            }
        }
        public bool countSpecified { get; set; }
        [XmlAttribute]
        public int uniqueCount
        {
            get
            {
                return this.uniqueCountField;
            }
            set
            {
                this.uniqueCountField = value;
                this.uniqueCountSpecified = true;
            }
        }
        public bool uniqueCountSpecified { get; set; }

    }
}
