using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;
using System.Xml;
using System.Xml.Schema;
using NPOI.OpenXml4Net.Util;
using System.IO;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot("pivotCacheDefinition", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public partial class CT_PivotCacheDefinition
    {
        public static CT_PivotCacheDefinition Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            var ctObj = new CT_PivotCacheDefinition();
            ctObj.id = XmlHelper.ReadString(node.Attributes[$"r:{nameof(id)}"]);
            if (node.Attributes[nameof(invalid)] != null)
                ctObj.invalid = XmlHelper.ReadBool(node.Attributes[nameof(invalid)]);
            if (node.Attributes[nameof(saveData)] != null)
                ctObj.saveData = XmlHelper.ReadBool(node.Attributes[nameof(saveData)]);
            if (node.Attributes[nameof(refreshOnLoad)] != null)
                ctObj.refreshOnLoad = XmlHelper.ReadBool(node.Attributes[nameof(refreshOnLoad)]);
            if (node.Attributes[nameof(optimizeMemory)] != null)
                ctObj.optimizeMemory = XmlHelper.ReadBool(node.Attributes[nameof(optimizeMemory)]);
            if (node.Attributes[nameof(enableRefresh)] != null)
                ctObj.enableRefresh = XmlHelper.ReadBool(node.Attributes[nameof(enableRefresh)]);
            ctObj.refreshedBy = XmlHelper.ReadString(node.Attributes[nameof(refreshedBy)]);
            if (node.Attributes[nameof(refreshedDate)] != null)
                ctObj.refreshedDate = XmlHelper.ReadDouble(node.Attributes[nameof(refreshedDate)]);
            if (node.Attributes[nameof(refreshedDateIso)] != null)
                if (node.Attributes[nameof(backgroundQuery)] != null)
                    ctObj.backgroundQuery = XmlHelper.ReadBool(node.Attributes[nameof(backgroundQuery)]);
            if (node.Attributes[nameof(missingItemsLimit)] != null)
                ctObj.missingItemsLimit = XmlHelper.ReadUInt(node.Attributes[nameof(missingItemsLimit)]);
            if (node.Attributes[nameof(createdVersion)] != null)
                ctObj.createdVersion = XmlHelper.ReadByte(node.Attributes[nameof(createdVersion)]);
            if (node.Attributes[nameof(refreshedVersion)] != null)
                ctObj.refreshedVersion = XmlHelper.ReadByte(node.Attributes[nameof(refreshedVersion)]);
            if (node.Attributes[nameof(minRefreshableVersion)] != null)
                ctObj.minRefreshableVersion = XmlHelper.ReadByte(node.Attributes[nameof(minRefreshableVersion)]);
            if (node.Attributes[nameof(recordCount)] != null)
                ctObj.recordCount = XmlHelper.ReadUInt(node.Attributes[nameof(recordCount)]);
            if (node.Attributes[nameof(upgradeOnRefresh)] != null)
                ctObj.upgradeOnRefresh = XmlHelper.ReadBool(node.Attributes[nameof(upgradeOnRefresh)]);
            if (node.Attributes[nameof(tupleCache1)] != null)
                ctObj.tupleCache1 = XmlHelper.ReadBool(node.Attributes[nameof(tupleCache1)]);
            if (node.Attributes[nameof(supportSubquery)] != null)
                ctObj.supportSubquery = XmlHelper.ReadBool(node.Attributes[nameof(supportSubquery)]);
            if (node.Attributes[nameof(supportAdvancedDrill)] != null)
                ctObj.supportAdvancedDrill = XmlHelper.ReadBool(node.Attributes[nameof(supportAdvancedDrill)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(cacheSource))
                    ctObj.cacheSource = CT_CacheSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(cacheFields))
                    ctObj.cacheFields = CT_CacheFields.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(cacheHierarchies))
                    ctObj.cacheHierarchies = CT_CacheHierarchies.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(kpis))
                    ctObj.kpis = CT_PCDKPIs.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(tupleCache))
                    ctObj.tupleCache = CT_TupleCache.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(calculatedItems))
                    ctObj.calculatedItems = CT_CalculatedItems.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(calculatedMembers))
                    ctObj.calculatedMembers = CT_CalculatedMembers.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(dimensions))
                    ctObj.dimensions = CT_Dimensions.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(measureGroups))
                    ctObj.measureGroups = CT_MeasureGroups.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(maps))
                    ctObj.maps = CT_MeasureDimensionMaps.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw)
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write("<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
            sw.Write("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ");
            sw.Write("xmlns:s=\"http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes\" ");
            XmlHelper.WriteAttribute(sw, $"r:{nameof(id)}", this.id);
            XmlHelper.WriteAttribute(sw, nameof(invalid), this.invalid);
            XmlHelper.WriteAttribute(sw, nameof(saveData), this.saveData);
            XmlHelper.WriteAttribute(sw, nameof(refreshOnLoad), this.refreshOnLoad);
            XmlHelper.WriteAttribute(sw, nameof(optimizeMemory), this.optimizeMemory);
            XmlHelper.WriteAttribute(sw, nameof(enableRefresh), this.enableRefresh);
            XmlHelper.WriteAttribute(sw, nameof(refreshedBy), this.refreshedBy);
            XmlHelper.WriteAttribute(sw, nameof(refreshedDate), this.refreshedDate);
            XmlHelper.WriteAttribute(sw, nameof(refreshedDateIso), this.refreshedDateIso);
            XmlHelper.WriteAttribute(sw, nameof(backgroundQuery), this.backgroundQuery);
            XmlHelper.WriteAttribute(sw, nameof(missingItemsLimit), this.missingItemsLimit);
            XmlHelper.WriteAttribute(sw, nameof(createdVersion), this.createdVersion);
            XmlHelper.WriteAttribute(sw, nameof(refreshedVersion), this.refreshedVersion);
            XmlHelper.WriteAttribute(sw, nameof(minRefreshableVersion), this.minRefreshableVersion);
            XmlHelper.WriteAttribute(sw, nameof(recordCount), this.recordCount);
            XmlHelper.WriteAttribute(sw, nameof(upgradeOnRefresh), this.upgradeOnRefresh);
            XmlHelper.WriteAttribute(sw, nameof(tupleCache1), this.tupleCache1);
            XmlHelper.WriteAttribute(sw, nameof(supportSubquery), this.supportSubquery);
            XmlHelper.WriteAttribute(sw, nameof(supportAdvancedDrill), this.supportAdvancedDrill);
            sw.Write(">");
            if (this.cacheSource != null)
                this.cacheSource.Write(sw, nameof(cacheSource));
            if (this.cacheFields != null)
                this.cacheFields.Write(sw, nameof(cacheFields));
            if (this.cacheHierarchies != null)
                this.cacheHierarchies.Write(sw, nameof(cacheHierarchies));
            if (this.kpis != null)
                this.kpis.Write(sw, nameof(kpis));
            if (this.tupleCache != null)
                this.tupleCache.Write(sw, nameof(tupleCache));
            if (this.calculatedItems != null)
                this.calculatedItems.Write(sw, nameof(calculatedItems));
            if (this.calculatedMembers != null)
                this.calculatedMembers.Write(sw, nameof(calculatedMembers));
            if (this.dimensions != null)
                this.dimensions.Write(sw, nameof(dimensions));
            if (this.measureGroups != null)
                this.measureGroups.Write(sw, nameof(measureGroups));
            if (this.maps != null)
                this.maps.Write(sw, nameof(maps));
            if (this.extLst != null)
                this.extLst.Write(sw, nameof(extLst));
            sw.Write(string.Format("</pivotCacheDefinition>"));
        }

        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                //TODO add namespaceUri
                this.Write(sw);
            }
        }

        public CT_PivotCacheDefinition()
        {
            this.extLst = new CT_ExtensionList();
            this.maps = new CT_MeasureDimensionMaps();
            this.measureGroups = new CT_MeasureGroups();
            this.dimensions = new CT_Dimensions();
            this.calculatedMembers = new CT_CalculatedMembers();
            this.calculatedItems = new CT_CalculatedItems();
            this.tupleCache = new CT_TupleCache();
            this.kpis = new CT_PCDKPIs();
            this.cacheHierarchies = new CT_CacheHierarchies();
            this.cacheFields = new CT_CacheFields();
            this.cacheSource = new CT_CacheSource();
            this.invalid = false;
            this.saveData = true;
            this.refreshOnLoad = false;
            this.optimizeMemory = false;
            this.enableRefresh = true;
            this.backgroundQuery = false;
            this.createdVersion = ((byte)(0));
            this.refreshedVersion = ((byte)(0));
            this.minRefreshableVersion = ((byte)(0));
            this.upgradeOnRefresh = false;
            this.tupleCache1 = false;
            this.supportSubquery = false;
            this.supportAdvancedDrill = false;
        }

        [XmlElement(Order = 0)]
        public CT_CacheSource cacheSource { get; set; }

        [XmlElement(Order = 1)]
        public CT_CacheFields cacheFields { get; set; }

        [XmlElement(Order = 2)]
        public CT_CacheHierarchies cacheHierarchies { get; set; }

        [XmlElement(Order = 3)]
        public CT_PCDKPIs kpis { get; set; }

        [XmlElement(Order = 4)]
        public CT_TupleCache tupleCache { get; set; }

        [XmlElement(Order = 5)]
        public CT_CalculatedItems calculatedItems { get; set; }

        [XmlElement(Order = 6)]
        public CT_CalculatedMembers calculatedMembers { get; set; }

        [XmlElement(Order = 7)]
        public CT_Dimensions dimensions { get; set; }

        [XmlElement(Order = 8)]
        public CT_MeasureGroups measureGroups { get; set; }

        [XmlElement(Order = 9)]
        public CT_MeasureDimensionMaps maps { get; set; }

        [XmlElement(Order = 10)]
        public CT_ExtensionList extLst { get; set; }

        [XmlAttribute(Form = XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool invalid { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool saveData { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool refreshOnLoad { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool optimizeMemory { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool enableRefresh { get; set; }

        [XmlAttribute()]
        public string refreshedBy { get; set; }

        [XmlAttribute()]
        public double refreshedDate { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool refreshedDateSpecified { get; set; }

        [XmlAttribute()]
        public System.DateTime? refreshedDateIso { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool refreshedDateIsoSpecified { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool backgroundQuery { get; set; }

        [XmlAttribute()]
        public uint missingItemsLimit { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool missingItemsLimitSpecified { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(byte), "0")]
        public byte createdVersion { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(byte), "0")]
        public byte refreshedVersion { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(byte), "0")]
        public byte minRefreshableVersion { get; set; }

        [XmlAttribute()]
        public uint recordCount { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool recordCountSpecified { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool upgradeOnRefresh { get; set; }

        [XmlAttribute(nameof(tupleCache))]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool tupleCache1 { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool supportSubquery { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool supportAdvancedDrill { get; set; }

        public CT_CacheFields AddNewCacheFields()
        {
            this.cacheFields = new CT_CacheFields();
            return this.cacheFields;
        }


        public CT_CacheSource AddNewCacheSource()
        {
            this.cacheSource = new CT_CacheSource();
            return this.cacheSource;
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Pages
    {
        public static CT_Pages Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Pages ctObj = new CT_Pages();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.page = new List<CT_PCDSCPage>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(page))
                    ctObj.page.Add(CT_PCDSCPage.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.page != null)
            {
                foreach (CT_PCDSCPage x in this.page)
                {
                    x.Write(sw, nameof(page));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Pages()
        {
            this.page = new List<CT_PCDSCPage>();
        }

        [XmlElement(nameof(page), Order = 0)]
        public List<CT_PCDSCPage> page { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PCDSCPage
    {
        public static CT_PCDSCPage Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PCDSCPage ctObj = new CT_PCDSCPage();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.pageItem = new List<CT_PageItem>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(pageItem))
                    ctObj.pageItem.Add(CT_PageItem.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.pageItem != null)
            {
                foreach (CT_PageItem x in this.pageItem)
                {
                    x.Write(sw, nameof(pageItem));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_PCDSCPage()
        {
            this.pageItem = new List<CT_PageItem>();
        }

        [XmlElement(nameof(pageItem), Order = 0)]
        public List<CT_PageItem> pageItem { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PageItem
    {
        public static CT_PageItem Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PageItem ctObj = new CT_PageItem();
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [XmlAttribute()]
        public string name { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_RangeSets
    {
        public static CT_RangeSets Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RangeSets ctObj = new CT_RangeSets();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.rangeSet = new List<CT_RangeSet>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(rangeSet))
                    ctObj.rangeSet.Add(CT_RangeSet.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.rangeSet != null)
            {
                foreach (CT_RangeSet x in this.rangeSet)
                {
                    x.Write(sw, nameof(rangeSet));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_RangeSets()
        {
            this.rangeSet = new List<CT_RangeSet>();
        }

        [XmlElement(nameof(rangeSet), Order = 0)]
        public List<CT_RangeSet> rangeSet { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_RangeSet
    {
        public static CT_RangeSet Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RangeSet ctObj = new CT_RangeSet();
            if (node.Attributes[nameof(i1)] != null)
                ctObj.i1 = XmlHelper.ReadUInt(node.Attributes[nameof(i1)]);
            if (node.Attributes[nameof(i2)] != null)
                ctObj.i2 = XmlHelper.ReadUInt(node.Attributes[nameof(i2)]);
            if (node.Attributes[nameof(i3)] != null)
                ctObj.i3 = XmlHelper.ReadUInt(node.Attributes[nameof(i3)]);
            if (node.Attributes[nameof(i4)] != null)
                ctObj.i4 = XmlHelper.ReadUInt(node.Attributes[nameof(i4)]);
            ctObj.@ref = XmlHelper.ReadString(node.Attributes[nameof(@ref)]);
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            ctObj.sheet = XmlHelper.ReadString(node.Attributes[nameof(sheet)]);
            ctObj.id = XmlHelper.ReadString(node.Attributes[$"r:{nameof(id)}"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(i1), this.i1);
            XmlHelper.WriteAttribute(sw, nameof(i2), this.i2);
            XmlHelper.WriteAttribute(sw, nameof(i3), this.i3);
            XmlHelper.WriteAttribute(sw, nameof(i4), this.i4);
            XmlHelper.WriteAttribute(sw, nameof(@ref), this.@ref);
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(sheet), this.sheet);
            XmlHelper.WriteAttribute(sw, $"r:{nameof(id)}", this.id);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [XmlAttribute()]
        public uint i1 { get; set; }

        [XmlIgnore()]
        public bool i1Specified { get; set; }

        [XmlAttribute()]
        public uint i2 { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool i2Specified { get; set; }

        [XmlAttribute()]
        public uint i3 { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool i3Specified { get; set; }

        [XmlAttribute()]
        public uint i4 { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool i4Specified { get; set; }

        [XmlAttribute()]
        public string @ref { get; set; }

        [XmlAttribute()]
        public string name { get; set; }

        [XmlAttribute()]
        public string sheet { get; set; }

        [XmlAttribute(Form = XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id { get; set; }
    }
    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Consolidation
    {
        public static CT_Consolidation Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Consolidation ctObj = new CT_Consolidation();
            if (node.Attributes[nameof(autoPage)] != null)
                ctObj.autoPage = XmlHelper.ReadBool(node.Attributes[nameof(autoPage)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(pages))
                    ctObj.pages = CT_Pages.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(rangeSets))
                    ctObj.rangeSets = CT_RangeSets.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(autoPage), this.autoPage);
            sw.Write(">");
            if (this.pages != null)
                this.pages.Write(sw, nameof(pages));
            if (this.rangeSets != null)
                this.rangeSets.Write(sw, nameof(rangeSets));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Consolidation()
        {
            this.rangeSets = new CT_RangeSets();
            this.pages = new CT_Pages();
            this.autoPage = true;
        }

        [XmlElement(Order = 0)]
        public CT_Pages pages { get; set; }

        [XmlElement(Order = 1)]
        public CT_RangeSets rangeSets { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool autoPage { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_WorksheetSource
    {
        public static CT_WorksheetSource Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_WorksheetSource ctObj = new CT_WorksheetSource();
            ctObj.@ref = XmlHelper.ReadString(node.Attributes[nameof(@ref)]);
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            ctObj.sheet = XmlHelper.ReadString(node.Attributes[nameof(sheet)]);
            ctObj.id = XmlHelper.ReadString(node.Attributes[$"r:{nameof(id)}"]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(@ref), this.@ref);
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(sheet), this.sheet);
            XmlHelper.WriteAttribute(sw, $"r:{nameof(id)}", this.id);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [XmlAttribute()]
        public string @ref { get; set; }

        [XmlAttribute()]
        public string name { get; set; }

        [XmlAttribute()]
        public string sheet { get; set; }

        [XmlAttribute(Form = XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id { get; set; }
    }

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_SourceType
    {

        /// <remarks/>
        worksheet,

        /// <remarks/>
        external,

        /// <remarks/>
        consolidation,

        /// <remarks/>
        scenario,
    }
    
    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_CacheSource
    {
        public static CT_CacheSource Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CacheSource ctObj = new CT_CacheSource();
            if (node.Attributes[nameof(type)] != null)
                ctObj.type = (ST_SourceType)Enum.Parse(typeof(ST_SourceType), node.Attributes[nameof(type)].Value);
            if (node.Attributes[nameof(connectionId)] != null)
                ctObj.connectionId = XmlHelper.ReadUInt(node.Attributes[nameof(connectionId)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(worksheetSource))
                    ctObj.worksheetSource = CT_WorksheetSource.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(consolidation))
                    ctObj.consolidation = CT_Consolidation.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(type), this.type.ToString());
            XmlHelper.WriteAttribute(sw, nameof(connectionId), this.connectionId);
            sw.Write(">");

            if (this.worksheetSource != null)
                this.worksheetSource.Write(sw, nameof(worksheetSource));
            if (this.consolidation != null)
                this.consolidation.Write(sw, nameof(consolidation));
            if (this.extLst != null)
                this.extLst.Write(sw, nameof(extLst));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_CacheSource()
        {
            this.connectionId = ((uint)(0));
        }

        [XmlElement("consolidation", typeof(CT_Consolidation), Order = 0)]
        [XmlElement("extLst", typeof(CT_ExtensionList), Order = 0)]
        [XmlElement("worksheetSource", typeof(CT_WorksheetSource), Order = 0)]
        public object Item { get; set; }

        [XmlAttribute()]
        public ST_SourceType type { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint connectionId { get; set; }

        public CT_WorksheetSource worksheetSource { get; set; }

        public CT_Consolidation consolidation { get; set; }

        public CT_ExtensionList extLst { get; set; }
        public CT_WorksheetSource AddNewWorksheetSource()
        {
            this.worksheetSource = new CT_WorksheetSource();
            return this.worksheetSource;
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_CacheFields
    {
        public static CT_CacheFields Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CacheFields ctObj = new CT_CacheFields();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.cacheField = new List<CT_CacheField>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(cacheField))
                    ctObj.cacheField.Add(CT_CacheField.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.cacheField != null)
            {
                foreach (CT_CacheField x in this.cacheField)
                {
                    x.Write(sw, nameof(cacheField));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_CacheFields()
        {
            this.cacheField = new List<CT_CacheField>();
        }

        [XmlElement(nameof(cacheField), Order = 0)]
        public List<CT_CacheField> cacheField { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }

        public CT_CacheField AddNewCacheField()
        {
            CT_CacheField f = new CT_CacheField();
            this.cacheField.Add(f);
            return f;
        }

        public uint SizeOfCacheFieldArray()
        {
            if (this.cacheField == null)
                return 0;
            return (uint)this.cacheField.Count;
        }
    }
    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_SharedItems
    {
        public static CT_SharedItems Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_SharedItems ctObj = new CT_SharedItems();
            if (node.Attributes[nameof(containsSemiMixedTypes)] != null)
                ctObj.containsSemiMixedTypes = XmlHelper.ReadBool(node.Attributes[nameof(containsSemiMixedTypes)]);
            if (node.Attributes[nameof(containsNonDate)] != null)
                ctObj.containsNonDate = XmlHelper.ReadBool(node.Attributes[nameof(containsNonDate)]);
            if (node.Attributes[nameof(containsDate)] != null)
                ctObj.containsDate = XmlHelper.ReadBool(node.Attributes[nameof(containsDate)]);
            if (node.Attributes[nameof(containsString)] != null)
                ctObj.containsString = XmlHelper.ReadBool(node.Attributes[nameof(containsString)]);
            if (node.Attributes[nameof(containsBlank)] != null)
                ctObj.containsBlank = XmlHelper.ReadBool(node.Attributes[nameof(containsBlank)]);
            if (node.Attributes[nameof(containsMixedTypes)] != null)
                ctObj.containsMixedTypes = XmlHelper.ReadBool(node.Attributes[nameof(containsMixedTypes)]);
            if (node.Attributes[nameof(containsNumber)] != null)
                ctObj.containsNumber = XmlHelper.ReadBool(node.Attributes[nameof(containsNumber)]);
            if (node.Attributes[nameof(containsInteger)] != null)
                ctObj.containsInteger = XmlHelper.ReadBool(node.Attributes[nameof(containsInteger)]);
            if (node.Attributes[nameof(minValue)] != null)
                ctObj.minValue = XmlHelper.ReadDouble(node.Attributes[nameof(minValue)]);
            if (node.Attributes[nameof(maxValue)] != null)
                ctObj.maxValue = XmlHelper.ReadDouble(node.Attributes[nameof(maxValue)]);
            if (node.Attributes[nameof(minDate)] != null)
                ctObj.minDate = XmlHelper.ReadDateTime(node.Attributes[nameof(minDate)]);
            if (node.Attributes[nameof(maxDate)] != null)
                ctObj.maxDate = XmlHelper.ReadDateTime(node.Attributes[nameof(maxDate)]);
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            if (node.Attributes[nameof(longText)] != null)
                ctObj.longText = XmlHelper.ReadBool(node.Attributes[nameof(longText)]);
            ctObj.Items = new List<Object>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "n")
                {
                    ctObj.Items.Add(CT_Number.Parse(childNode, namespaceManager));
                    //ctObj.ItemsElementName.Add(ItemsChoiceType1.n);
                }
                else if (childNode.LocalName == "b")
                {
                    ctObj.Items.Add(CT_Boolean.Parse(childNode, namespaceManager));
                    //ctObj.ItemsElementName.Add(ItemsChoiceType1.b);
                }
                else if (childNode.LocalName == "d")
                {
                    ctObj.Items.Add(CT_DateTime.Parse(childNode, namespaceManager));
                    //ctObj.ItemsElementName.Add(ItemsChoiceType1.d);
                }
                else if (childNode.LocalName == "e")
                {
                    ctObj.Items.Add(CT_Error.Parse(childNode, namespaceManager));
                    //ctObj.ItemsElementName.Add(ItemsChoiceType1.e);
                }
                else if (childNode.LocalName == "m")
                {
                    ctObj.Items.Add(CT_Missing.Parse(childNode, namespaceManager));
                    //ctObj.ItemsElementName.Add(ItemsChoiceType1.m);
                }
                else if (childNode.LocalName == "s")
                {
                    ctObj.Items.Add(CT_String.Parse(childNode, namespaceManager));
                    //ctObj.ItemsElementName.Add(ItemsChoiceType1.s);
                }
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(containsSemiMixedTypes), this.containsSemiMixedTypes);
            XmlHelper.WriteAttribute(sw, nameof(containsNonDate), this.containsNonDate);
            XmlHelper.WriteAttribute(sw, nameof(containsDate), this.containsDate);
            XmlHelper.WriteAttribute(sw, nameof(containsString), this.containsString);
            XmlHelper.WriteAttribute(sw, nameof(containsBlank), this.containsBlank);
            XmlHelper.WriteAttribute(sw, nameof(containsMixedTypes), this.containsMixedTypes);
            XmlHelper.WriteAttribute(sw, nameof(containsNumber), this.containsNumber);
            XmlHelper.WriteAttribute(sw, nameof(containsInteger), this.containsInteger);
            XmlHelper.WriteAttribute(sw, nameof(minValue), this.minValue);
            XmlHelper.WriteAttribute(sw, nameof(maxValue), this.maxValue);
            XmlHelper.WriteAttribute(sw, nameof(minDate), this.minDate);
            XmlHelper.WriteAttribute(sw, nameof(maxDate), this.maxDate);
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            XmlHelper.WriteAttribute(sw, nameof(longText), this.longText);
            sw.Write(">");

            foreach (object o in this.Items)
            {
                if (o is CT_Number)
                    ((CT_Number)o).Write(sw, "n");
                else if (o is CT_Boolean)
                    ((CT_Boolean)o).Write(sw, "b");
                else if (o is CT_DateTime)
                    ((CT_DateTime)o).Write(sw, "d");
                else if (o is CT_Error)
                    ((CT_Error)o).Write(sw, "e");
                else if (o is CT_Missing)
                    ((CT_Missing)o).Write(sw, "m");
                else if (o is CT_String)
                    ((CT_String)o).Write(sw, "s");
            }

            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_SharedItems()
        {
            this.Items = new List<object>();
            this.containsSemiMixedTypes = true;
            this.containsNonDate = true;
            this.containsDate = false;
            this.containsString = true;
            this.containsBlank = false;
            this.containsMixedTypes = false;
            this.containsNumber = false;
            this.containsInteger = false;
            this.longText = false;
        }

        [XmlElement("b", typeof(CT_Boolean), Order = 0)]
        [XmlElement("d", typeof(CT_DateTime), Order = 0)]
        [XmlElement("e", typeof(CT_Error), Order = 0)]
        [XmlElement("m", typeof(CT_Missing), Order = 0)]
        [XmlElement("n", typeof(CT_Number), Order = 0)]
        [XmlElement("s", typeof(CT_String), Order = 0)]
        public List<object> Items { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool containsSemiMixedTypes { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool containsNonDate { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool containsDate { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool containsString { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool containsBlank { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool containsMixedTypes { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool containsNumber { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool containsInteger { get; set; }

        [XmlAttribute()]
        public double minValue { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool minValueSpecified { get; set; }

        [XmlAttribute()]
        public double maxValue { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool maxValueSpecified { get; set; }

        [XmlAttribute()]
        public System.DateTime? minDate { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool minDateSpecified { get; set; }

        [XmlAttribute()]
        public System.DateTime? maxDate { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool maxDateSpecified { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool longText { get; set; }
    }
    
    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_CacheField
    {
        public static CT_CacheField Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CacheField ctObj = new CT_CacheField();
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            ctObj.caption = XmlHelper.ReadString(node.Attributes[nameof(caption)]);
            ctObj.propertyName = XmlHelper.ReadString(node.Attributes[nameof(propertyName)]);
            if (node.Attributes[nameof(serverField)] != null)
                ctObj.serverField = XmlHelper.ReadBool(node.Attributes[nameof(serverField)]);
            if (node.Attributes[nameof(uniqueList)] != null)
                ctObj.uniqueList = XmlHelper.ReadBool(node.Attributes[nameof(uniqueList)]);
            if (node.Attributes[nameof(numFmtId)] != null)
                ctObj.numFmtId = XmlHelper.ReadUInt(node.Attributes[nameof(numFmtId)]);
            ctObj.formula = XmlHelper.ReadString(node.Attributes[nameof(formula)]);
            if (node.Attributes[nameof(sqlType)] != null)
                ctObj.sqlType = XmlHelper.ReadInt(node.Attributes[nameof(sqlType)]);
            if (node.Attributes[nameof(hierarchy)] != null)
                ctObj.hierarchy = XmlHelper.ReadInt(node.Attributes[nameof(hierarchy)]);
            if (node.Attributes[nameof(level)] != null)
                ctObj.level = XmlHelper.ReadUInt(node.Attributes[nameof(level)]);
            if (node.Attributes[nameof(databaseField)] != null)
                ctObj.databaseField = XmlHelper.ReadBool(node.Attributes[nameof(databaseField)]);
            if (node.Attributes[nameof(mappingCount)] != null)
                ctObj.mappingCount = XmlHelper.ReadUInt(node.Attributes[nameof(mappingCount)]);
            if (node.Attributes[nameof(memberPropertyField)] != null)
                ctObj.memberPropertyField = XmlHelper.ReadBool(node.Attributes[nameof(memberPropertyField)]);
            ctObj.mpMap = new List<CT_X>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(sharedItems))
                    ctObj.sharedItems = CT_SharedItems.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(fieldGroup))
                    ctObj.fieldGroup = CT_FieldGroup.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(mpMap))
                    ctObj.mpMap.Add(CT_X.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(caption), this.caption);
            XmlHelper.WriteAttribute(sw, nameof(propertyName), this.propertyName);
            XmlHelper.WriteAttribute(sw, nameof(serverField), this.serverField);
            XmlHelper.WriteAttribute(sw, nameof(uniqueList), this.uniqueList);
            XmlHelper.WriteAttribute(sw, nameof(numFmtId), this.numFmtId);
            XmlHelper.WriteAttribute(sw, nameof(formula), this.formula);
            XmlHelper.WriteAttribute(sw, nameof(sqlType), this.sqlType);
            XmlHelper.WriteAttribute(sw, nameof(hierarchy), this.hierarchy);
            XmlHelper.WriteAttribute(sw, nameof(level), this.level);
            XmlHelper.WriteAttribute(sw, nameof(databaseField), this.databaseField);
            XmlHelper.WriteAttribute(sw, nameof(mappingCount), this.mappingCount);
            XmlHelper.WriteAttribute(sw, nameof(memberPropertyField), this.memberPropertyField);
            sw.Write(">");
            if (this.sharedItems != null)
                this.sharedItems.Write(sw, nameof(sharedItems));
            if (this.fieldGroup != null)
                this.fieldGroup.Write(sw, nameof(fieldGroup));
            if (this.extLst != null)
                this.extLst.Write(sw, nameof(extLst));
            if (this.mpMap != null)
            {
                foreach (CT_X x in this.mpMap)
                {
                    x.Write(sw, nameof(mpMap));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_CacheField()
        {
            this.extLst = new CT_ExtensionList();
            this.mpMap = new List<CT_X>();
            this.fieldGroup = new CT_FieldGroup();
            this.sharedItems = new CT_SharedItems();
            this.serverField = false;
            this.uniqueList = true;
            this.sqlType = 0;
            this.hierarchy = 0;
            this.level = ((uint)(0));
            this.databaseField = true;
            this.memberPropertyField = false;
        }

        [XmlElement(Order = 0)]
        public CT_SharedItems sharedItems { get; set; }

        [XmlElement(Order = 1)]
        public CT_FieldGroup fieldGroup { get; set; }

        [XmlElement(nameof(mpMap), Order = 2)]
        public List<CT_X> mpMap { get; set; }

        [XmlElement(Order = 3)]
        public CT_ExtensionList extLst { get; set; }

        [XmlAttribute()]
        public string name { get; set; }

        [XmlAttribute()]
        public string caption { get; set; }

        [XmlAttribute()]
        public string propertyName { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool serverField { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool uniqueList { get; set; }

        [XmlAttribute()]
        public uint numFmtId { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool numFmtIdSpecified { get; set; }

        [XmlAttribute()]
        public string formula { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public int sqlType { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public int hierarchy { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint level { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool databaseField { get; set; }

        [XmlAttribute()]
        public uint mappingCount { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool mappingCountSpecified { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool memberPropertyField { get; set; }

        public CT_SharedItems AddNewSharedItems()
        {
            this.sharedItems = new CT_SharedItems();
            return this.sharedItems;
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_CacheHierarchies
    {
        public static CT_CacheHierarchies Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CacheHierarchies ctObj = new CT_CacheHierarchies();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.cacheHierarchy = new List<CT_CacheHierarchy>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(cacheHierarchy))
                    ctObj.cacheHierarchy.Add(CT_CacheHierarchy.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.cacheHierarchy != null)
            {
                foreach (CT_CacheHierarchy x in this.cacheHierarchy)
                {
                    x.Write(sw, nameof(cacheHierarchy));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_CacheHierarchies()
        {
            this.cacheHierarchy = new List<CT_CacheHierarchy>();
        }

        [XmlElement(nameof(cacheHierarchy), Order = 0)]
        public List<CT_CacheHierarchy> cacheHierarchy { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Boolean
    {
        public static CT_Boolean Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Boolean ctObj = new CT_Boolean();
            if (node.Attributes[nameof(v)] != null)
                ctObj.v = XmlHelper.ReadBool(node.Attributes[nameof(v)]);
            if (node.Attributes[nameof(u)] != null)
                ctObj.u = XmlHelper.ReadBool(node.Attributes[nameof(u)]);
            if (node.Attributes[nameof(f)] != null)
                ctObj.f = XmlHelper.ReadBool(node.Attributes[nameof(f)]);
            ctObj.c = XmlHelper.ReadString(node.Attributes[nameof(c)]);
            if (node.Attributes[nameof(cp)] != null)
                ctObj.cp = XmlHelper.ReadUInt(node.Attributes[nameof(cp)]);
            ctObj.x = new List<CT_X>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(x))
                    ctObj.x.Add(CT_X.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(v), this.v);
            XmlHelper.WriteAttribute(sw, nameof(u), this.u);
            XmlHelper.WriteAttribute(sw, nameof(f), this.f);
            XmlHelper.WriteAttribute(sw, nameof(c), this.c);
            XmlHelper.WriteAttribute(sw, nameof(cp), this.cp);
            sw.Write(">");
            if (this.x != null)
            {
                foreach (CT_X x in this.x)
                {
                    x.Write(sw, nameof(x));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Boolean()
        {
            this.x = new List<CT_X>();
        }

        [XmlElement(nameof(x), Order = 0)]
        public List<CT_X> x { get; set; }

        [XmlAttribute()]
        public bool v { get; set; }

        [XmlAttribute()]
        public bool u { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool uSpecified { get; set; }

        [XmlAttribute()]
        public bool f { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool fSpecified { get; set; }

        [XmlAttribute()]
        public string c { get; set; }

        [XmlAttribute()]
        public uint cp { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool cpSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_X
    {
        public CT_X()
        {
            this.v = 0;
        }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public int v { get; set; }

        public static CT_X Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_X ctObj = new CT_X();
            if (node.Attributes[nameof(v)] != null)
                ctObj.v = XmlHelper.ReadInt(node.Attributes[nameof(v)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(v), this.v);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_DateTime
    {
        public static CT_DateTime Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_DateTime ctObj = new CT_DateTime();
            if (node.Attributes[nameof(v)] != null)
                ctObj.v = XmlHelper.ReadDateTime(node.Attributes[nameof(v)]); 
            if (node.Attributes[nameof(u)] != null)
                ctObj.u = XmlHelper.ReadBool(node.Attributes[nameof(u)]);
            if (node.Attributes[nameof(f)] != null)
                ctObj.f = XmlHelper.ReadBool(node.Attributes[nameof(f)]);
            ctObj.c = XmlHelper.ReadString(node.Attributes[nameof(c)]);
            if (node.Attributes[nameof(cp)] != null)
                ctObj.cp = XmlHelper.ReadUInt(node.Attributes[nameof(cp)]);
            ctObj.x = new List<CT_X>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(x))
                    ctObj.x.Add(CT_X.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(v), this.v);
            XmlHelper.WriteAttribute(sw, nameof(u), this.u);
            XmlHelper.WriteAttribute(sw, nameof(f), this.f);
            XmlHelper.WriteAttribute(sw, nameof(c), this.c);
            XmlHelper.WriteAttribute(sw, nameof(cp), this.cp);
            sw.Write(">");
            if (this.x != null)
            {
                foreach (CT_X x in this.x)
                {
                    x.Write(sw, nameof(x));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_DateTime()
        {
            this.x = new List<CT_X>();
        }

        [XmlElement(nameof(x), Order = 0)]
        public List<CT_X> x { get; set; }

        [XmlAttribute()]
        public System.DateTime? v { get; set; }

        [XmlAttribute()]
        public bool u { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool uSpecified { get; set; }

        [XmlAttribute()]
        public bool f { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool fSpecified { get; set; }

        [XmlAttribute()]
        public string c { get; set; }

        [XmlAttribute()]
        public uint cp { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool cpSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Error
    {
        public CT_Error()
        {
            this.x = new List<CT_X>();
            this.tpls = new CT_Tuples();
            this.i = false;
            this.un = false;
            this.st = false;
            this.b = false;
        }

        [XmlElement(Order = 0)]
        public CT_Tuples tpls { get; set; }

        [XmlElement(nameof(x), Order = 1)]
        public List<CT_X> x { get; set; }

        [XmlAttribute()]
        public string v { get; set; }

        [XmlAttribute()]
        public bool u { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool uSpecified { get; set; }

        [XmlAttribute()]
        public bool f { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool fSpecified { get; set; }

        [XmlAttribute()]
        public string c { get; set; }

        [XmlAttribute()]
        public uint cp { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool cpSpecified { get; set; }

        [XmlAttribute()]
        public uint @in { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool inSpecified { get; set; }

        [XmlAttribute(DataType = "hexBinary")]
        public byte[] bc { get; set; }

        [XmlAttribute(DataType = "hexBinary")]
        public byte[] fc { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool i { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool un { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool st { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool b { get; set; }

        public static CT_Error Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Error ctObj = new CT_Error();
            ctObj.v = XmlHelper.ReadString(node.Attributes[nameof(v)]);
            if (node.Attributes[nameof(u)] != null)
                ctObj.u = XmlHelper.ReadBool(node.Attributes[nameof(u)]);
            if (node.Attributes[nameof(f)] != null)
                ctObj.f = XmlHelper.ReadBool(node.Attributes[nameof(f)]);
            ctObj.c = XmlHelper.ReadString(node.Attributes[nameof(c)]);
            if (node.Attributes[nameof(cp)] != null)
                ctObj.cp = XmlHelper.ReadUInt(node.Attributes[nameof(cp)]);
            if (node.Attributes[nameof(@in)] != null)
                ctObj.@in = XmlHelper.ReadUInt(node.Attributes[nameof(@in)]);
            ctObj.bc = XmlHelper.ReadBytes(node.Attributes[nameof(bc)]);
            ctObj.fc = XmlHelper.ReadBytes(node.Attributes[nameof(fc)]);
            if (node.Attributes[nameof(i)] != null)
                ctObj.i = XmlHelper.ReadBool(node.Attributes[nameof(i)]);
            if (node.Attributes[nameof(un)] != null)
                ctObj.un = XmlHelper.ReadBool(node.Attributes[nameof(un)]);
            if (node.Attributes[nameof(st)] != null)
                ctObj.st = XmlHelper.ReadBool(node.Attributes[nameof(st)]);
            if (node.Attributes[nameof(b)] != null)
                ctObj.b = XmlHelper.ReadBool(node.Attributes[nameof(b)]);
            ctObj.x = new List<CT_X>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(tpls))
                    ctObj.tpls = CT_Tuples.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(x))
                    ctObj.x.Add(CT_X.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(v), this.v);
            XmlHelper.WriteAttribute(sw, nameof(u), this.u);
            XmlHelper.WriteAttribute(sw, nameof(f), this.f);
            XmlHelper.WriteAttribute(sw, nameof(c), this.c);
            XmlHelper.WriteAttribute(sw, nameof(cp), this.cp);
            XmlHelper.WriteAttribute(sw, nameof(@in), this.@in);
            XmlHelper.WriteAttribute(sw, nameof(bc), this.bc);
            XmlHelper.WriteAttribute(sw, nameof(fc), this.fc);
            XmlHelper.WriteAttribute(sw, nameof(i), this.i);
            XmlHelper.WriteAttribute(sw, nameof(un), this.un);
            XmlHelper.WriteAttribute(sw, nameof(st), this.st);
            XmlHelper.WriteAttribute(sw, nameof(b), this.b);
            sw.Write(">");
            if (this.tpls != null)
                this.tpls.Write(sw, nameof(tpls));
            if (this.x != null)
            {
                foreach (CT_X x in this.x)
                {
                    x.Write(sw, nameof(x));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Tuples
    {
        public CT_Tuples()
        {
            this.tpl = new List<CT_Tuple>();
        }

        [XmlElement(nameof(tpl), Order = 0)]
        public List<CT_Tuple> tpl { get; set; }

        [XmlAttribute()]
        public uint c { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool cSpecified { get; set; }

        public static CT_Tuples Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Tuples ctObj = new CT_Tuples();
            if (node.Attributes[nameof(c)] != null)
                ctObj.c = XmlHelper.ReadUInt(node.Attributes[nameof(c)]);
            ctObj.tpl = new List<CT_Tuple>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(tpl))
                    ctObj.tpl.Add(CT_Tuple.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(c), this.c);
            sw.Write(">");
            if (this.tpl != null)
            {
                foreach (CT_Tuple x in this.tpl)
                {
                    x.Write(sw, nameof(tpl));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Tuple
    {
        [XmlAttribute()]
        public uint fld { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool fldSpecified { get; set; }

        [XmlAttribute()]
        public uint hier { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool hierSpecified { get; set; }

        [XmlAttribute()]
        public uint item { get; set; }

        public static CT_Tuple Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Tuple ctObj = new CT_Tuple();
            if (node.Attributes[nameof(fld)] != null)
                ctObj.fld = XmlHelper.ReadUInt(node.Attributes[nameof(fld)]);
            if (node.Attributes[nameof(hier)] != null)
                ctObj.hier = XmlHelper.ReadUInt(node.Attributes[nameof(hier)]);
            if (node.Attributes[nameof(item)] != null)
                ctObj.item = XmlHelper.ReadUInt(node.Attributes[nameof(item)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(fld), this.fld);
            XmlHelper.WriteAttribute(sw, nameof(hier), this.hier);
            XmlHelper.WriteAttribute(sw, nameof(item), this.item);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

    }




    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Missing
    {
        public CT_Missing()
        {
            this.x = new List<CT_X>();
            this.tpls = new List<CT_Tuples>();
            this.i = false;
            this.un = false;
            this.st = false;
            this.b = false;
        }

        [XmlElement(nameof(tpls), Order = 0)]
        public List<CT_Tuples> tpls { get; set; }

        [XmlElement(nameof(x), Order = 1)]
        public List<CT_X> x { get; set; }

        [XmlAttribute()]
        public bool u { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool uSpecified { get; set; }

        [XmlAttribute()]
        public bool f { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool fSpecified { get; set; }

        [XmlAttribute()]
        public string c { get; set; }

        [XmlAttribute()]
        public uint cp { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool cpSpecified { get; set; }

        [XmlAttribute()]
        public uint @in { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool inSpecified { get; set; }

        [XmlAttribute(DataType = "hexBinary")]
        public byte[] bc { get; set; }

        [XmlAttribute(DataType = "hexBinary")]
        public byte[] fc { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool i { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool un { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool st { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool b { get; set; }

        public static CT_Missing Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Missing ctObj = new CT_Missing();
            if (node.Attributes[nameof(u)] != null)
                ctObj.u = XmlHelper.ReadBool(node.Attributes[nameof(u)]);
            if (node.Attributes[nameof(f)] != null)
                ctObj.f = XmlHelper.ReadBool(node.Attributes[nameof(f)]);
            ctObj.c = XmlHelper.ReadString(node.Attributes[nameof(c)]);
            if (node.Attributes[nameof(cp)] != null)
                ctObj.cp = XmlHelper.ReadUInt(node.Attributes[nameof(cp)]);
            if (node.Attributes[nameof(@in)] != null)
                ctObj.@in = XmlHelper.ReadUInt(node.Attributes[nameof(@in)]);
            ctObj.bc = XmlHelper.ReadBytes(node.Attributes[nameof(bc)]);
            ctObj.fc = XmlHelper.ReadBytes(node.Attributes[nameof(fc)]);
            if (node.Attributes[nameof(i)] != null)
                ctObj.i = XmlHelper.ReadBool(node.Attributes[nameof(i)]);
            if (node.Attributes[nameof(un)] != null)
                ctObj.un = XmlHelper.ReadBool(node.Attributes[nameof(un)]);
            if (node.Attributes[nameof(st)] != null)
                ctObj.st = XmlHelper.ReadBool(node.Attributes[nameof(st)]);
            if (node.Attributes[nameof(b)] != null)
                ctObj.b = XmlHelper.ReadBool(node.Attributes[nameof(b)]);
            ctObj.tpls = new List<CT_Tuples>();
            ctObj.x = new List<CT_X>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(tpls))
                    ctObj.tpls.Add(CT_Tuples.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == nameof(x))
                    ctObj.x.Add(CT_X.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(u), this.u);
            XmlHelper.WriteAttribute(sw, nameof(f), this.f);
            XmlHelper.WriteAttribute(sw, nameof(c), this.c);
            XmlHelper.WriteAttribute(sw, nameof(cp), this.cp);
            XmlHelper.WriteAttribute(sw, nameof(@in), this.@in);
            XmlHelper.WriteAttribute(sw, nameof(bc), this.bc);
            XmlHelper.WriteAttribute(sw, nameof(fc), this.fc);
            XmlHelper.WriteAttribute(sw, nameof(i), this.i);
            XmlHelper.WriteAttribute(sw, nameof(un), this.un);
            XmlHelper.WriteAttribute(sw, nameof(st), this.st);
            XmlHelper.WriteAttribute(sw, nameof(b), this.b);
            sw.Write(">");
            if (this.tpls != null)
            {
                foreach (CT_Tuples x in this.tpls)
                {
                    x.Write(sw, nameof(tpls));
                }
            }
            if (this.x != null)
            {
                foreach (CT_X x in this.x)
                {
                    x.Write(sw, "x");
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Number
    {
        public CT_Number()
        {
            this.x = new List<CT_X>();
            this.tpls = new List<CT_Tuples>();
            this.i = false;
            this.un = false;
            this.st = false;
            this.b = false;
        }

        [XmlElement(nameof(tpls), Order = 0)]
        public List<CT_Tuples> tpls { get; set; }

        [XmlElement(nameof(x), Order = 1)]
        public List<CT_X> x { get; set; }

        [XmlAttribute()]
        public double v { get; set; }

        [XmlAttribute()]
        public bool u { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool uSpecified { get; set; }

        [XmlAttribute()]
        public bool f { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool fSpecified { get; set; }

        [XmlAttribute()]
        public string c { get; set; }

        [XmlAttribute()]
        public uint cp { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool cpSpecified { get; set; }

        [XmlAttribute()]
        public uint @in { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool inSpecified { get; set; }

        [XmlAttribute(DataType = "hexBinary")]
        public byte[] bc { get; set; }

        [XmlAttribute(DataType = "hexBinary")]
        public byte[] fc { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool i { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool un { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool st { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool b { get; set; }

        public static CT_Number Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Number ctObj = new CT_Number();
            if (node.Attributes[nameof(v)] != null)
                ctObj.v = XmlHelper.ReadDouble(node.Attributes[nameof(v)]);
            if (node.Attributes[nameof(u)] != null)
                ctObj.u = XmlHelper.ReadBool(node.Attributes[nameof(u)]);
            if (node.Attributes[nameof(f)] != null)
                ctObj.f = XmlHelper.ReadBool(node.Attributes[nameof(f)]);
            ctObj.c = XmlHelper.ReadString(node.Attributes[nameof(c)]);
            if (node.Attributes[nameof(cp)] != null)
                ctObj.cp = XmlHelper.ReadUInt(node.Attributes[nameof(cp)]);
            if (node.Attributes[nameof(@in)] != null)
                ctObj.@in = XmlHelper.ReadUInt(node.Attributes[nameof(@in)]);
            ctObj.bc = XmlHelper.ReadBytes(node.Attributes[nameof(bc)]);
            ctObj.fc = XmlHelper.ReadBytes(node.Attributes[nameof(fc)]);
            if (node.Attributes[nameof(i)] != null)
                ctObj.i = XmlHelper.ReadBool(node.Attributes[nameof(i)]);
            if (node.Attributes[nameof(un)] != null)
                ctObj.un = XmlHelper.ReadBool(node.Attributes[nameof(un)]);
            if (node.Attributes[nameof(st)] != null)
                ctObj.st = XmlHelper.ReadBool(node.Attributes[nameof(st)]);
            if (node.Attributes[nameof(b)] != null)
                ctObj.b = XmlHelper.ReadBool(node.Attributes[nameof(b)]);
            ctObj.tpls = new List<CT_Tuples>();
            ctObj.x = new List<CT_X>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(tpls))
                    ctObj.tpls.Add(CT_Tuples.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == nameof(x))
                    ctObj.x.Add(CT_X.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(v), this.v);
            XmlHelper.WriteAttribute(sw, nameof(u), this.u);
            XmlHelper.WriteAttribute(sw, nameof(f), this.f);
            XmlHelper.WriteAttribute(sw, nameof(c), this.c);
            XmlHelper.WriteAttribute(sw, nameof(cp), this.cp);
            XmlHelper.WriteAttribute(sw, nameof(@in), this.@in);
            XmlHelper.WriteAttribute(sw, nameof(bc), this.bc);
            XmlHelper.WriteAttribute(sw, nameof(fc), this.fc);
            XmlHelper.WriteAttribute(sw, nameof(i), this.i);
            XmlHelper.WriteAttribute(sw, nameof(un), this.un);
            XmlHelper.WriteAttribute(sw, nameof(st), this.st);
            XmlHelper.WriteAttribute(sw, nameof(b), this.b);
            sw.Write(">");
            if (this.tpls != null)
            {
                foreach (CT_Tuples x in this.tpls)
                {
                    x.Write(sw, nameof(tpls));
                }
            }
            if (this.x != null)
            {
                foreach (CT_X x in this.x)
                {
                    x.Write(sw, nameof(x));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_String
    {
        public CT_String()
        {
            this.x = new List<CT_X>();
            this.tpls = new List<CT_Tuples>();
            this.i = false;
            this.un = false;
            this.st = false;
            this.b = false;
        }

        [XmlElement(nameof(tpls), Order = 0)]
        public List<CT_Tuples> tpls { get; set; }

        [XmlElement(nameof(x), Order = 1)]
        public List<CT_X> x { get; set; }

        [XmlAttribute()]
        public string v { get; set; }

        [XmlAttribute()]
        public bool u { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool uSpecified { get; set; }

        [XmlAttribute()]
        public bool f { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool fSpecified { get; set; }

        [XmlAttribute()]
        public string c { get; set; }

        [XmlAttribute()]
        public uint cp { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool cpSpecified { get; set; }

        [XmlAttribute()]
        public uint @in { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool inSpecified { get; set; }

        [XmlAttribute(DataType = "hexBinary")]
        public byte[] bc { get; set; }

        [XmlAttribute(DataType = "hexBinary")]
        public byte[] fc { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool i { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool un { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool st { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool b { get; set; }

        public static CT_String Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_String ctObj = new CT_String();
            ctObj.v = XmlHelper.ReadString(node.Attributes[nameof(v)]);
            if (node.Attributes[nameof(u)] != null)
                ctObj.u = XmlHelper.ReadBool(node.Attributes[nameof(u)]);
            if (node.Attributes[nameof(f)] != null)
                ctObj.f = XmlHelper.ReadBool(node.Attributes[nameof(f)]);
            ctObj.c = XmlHelper.ReadString(node.Attributes[nameof(c)]);
            if (node.Attributes[nameof(cp)] != null)
                ctObj.cp = XmlHelper.ReadUInt(node.Attributes[nameof(cp)]);
            if (node.Attributes[nameof(@in)] != null)
                ctObj.@in = XmlHelper.ReadUInt(node.Attributes[nameof(@in)]);
            ctObj.bc = XmlHelper.ReadBytes(node.Attributes[nameof(bc)]);
            ctObj.fc = XmlHelper.ReadBytes(node.Attributes[nameof(fc)]);
            if (node.Attributes[nameof(i)] != null)
                ctObj.i = XmlHelper.ReadBool(node.Attributes[nameof(i)]);
            if (node.Attributes[nameof(un)] != null)
                ctObj.un = XmlHelper.ReadBool(node.Attributes[nameof(un)]);
            if (node.Attributes[nameof(st)] != null)
                ctObj.st = XmlHelper.ReadBool(node.Attributes[nameof(st)]);
            if (node.Attributes[nameof(b)] != null)
                ctObj.b = XmlHelper.ReadBool(node.Attributes[nameof(b)]);
            ctObj.tpls = new List<CT_Tuples>();
            ctObj.x = new List<CT_X>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(tpls))
                    ctObj.tpls.Add(CT_Tuples.Parse(childNode, namespaceManager));
                else if (childNode.LocalName == nameof(x))
                    ctObj.x.Add(CT_X.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(v), this.v);
            XmlHelper.WriteAttribute(sw, nameof(u), this.u);
            XmlHelper.WriteAttribute(sw, nameof(f), this.f);
            XmlHelper.WriteAttribute(sw, nameof(c), this.c);
            XmlHelper.WriteAttribute(sw, nameof(cp), this.cp);
            XmlHelper.WriteAttribute(sw, nameof(@in), this.@in);
            XmlHelper.WriteAttribute(sw, nameof(bc), this.bc);
            XmlHelper.WriteAttribute(sw, nameof(fc), this.fc);
            XmlHelper.WriteAttribute(sw, nameof(i), this.i);
            XmlHelper.WriteAttribute(sw, nameof(un), this.un);
            XmlHelper.WriteAttribute(sw, nameof(st), this.st);
            XmlHelper.WriteAttribute(sw, nameof(b), this.b);
            sw.Write(">");
            if (this.tpls != null)
            {
                foreach (CT_Tuples x in this.tpls)
                {
                    x.Write(sw, nameof(tpls));
                }
            }
            if (this.x != null)
            {
                foreach (CT_X x in this.x)
                {
                    x.Write(sw, nameof(x));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_FieldGroup
    {
        public static CT_FieldGroup Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FieldGroup ctObj = new CT_FieldGroup();
            if (node.Attributes[nameof(par)] != null)
                ctObj.par = XmlHelper.ReadUInt(node.Attributes[nameof(par)]);
            if (node.Attributes[nameof(@base)] != null)
                ctObj.@base = XmlHelper.ReadUInt(node.Attributes[nameof(@base)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(rangePr))
                    ctObj.rangePr = CT_RangePr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(discretePr))
                    ctObj.discretePr = CT_DiscretePr.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(groupItems))
                    ctObj.groupItems = CT_GroupItems.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(par), this.par);
            XmlHelper.WriteAttribute(sw, nameof(@base), this.@base);
            sw.Write(">");
            if (this.rangePr != null)
                this.rangePr.Write(sw, nameof(rangePr));
            if (this.discretePr != null)
                this.discretePr.Write(sw, nameof(discretePr));
            if (this.groupItems != null)
                this.groupItems.Write(sw, nameof(groupItems));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_FieldGroup()
        {
            this.groupItems = new CT_GroupItems();
            this.discretePr = new CT_DiscretePr();
            this.rangePr = new CT_RangePr();
        }

        [XmlElement(Order = 0)]
        public CT_RangePr rangePr { get; set; }

        [XmlElement(Order = 1)]
        public CT_DiscretePr discretePr { get; set; }

        [XmlElement(Order = 2)]
        public CT_GroupItems groupItems { get; set; }

        [XmlAttribute()]
        public uint par { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool parSpecified { get; set; }

        [XmlAttribute()]
        public uint @base { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool baseSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_RangePr
    {
        public static CT_RangePr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RangePr ctObj = new CT_RangePr();
            if (node.Attributes[nameof(autoStart)] != null)
                ctObj.autoStart = XmlHelper.ReadBool(node.Attributes[nameof(autoStart)]);
            if (node.Attributes[nameof(autoEnd)] != null)
                ctObj.autoEnd = XmlHelper.ReadBool(node.Attributes[nameof(autoEnd)]);
            if (node.Attributes[nameof(groupBy)] != null)
                ctObj.groupBy = (ST_GroupBy)Enum.Parse(typeof(ST_GroupBy), node.Attributes[nameof(groupBy)].Value);
            if (node.Attributes[nameof(startNum)] != null)
                ctObj.startNum = XmlHelper.ReadDouble(node.Attributes[nameof(startNum)]);
            if (node.Attributes[nameof(endNum)] != null)
                ctObj.endNum = XmlHelper.ReadDouble(node.Attributes[nameof(endNum)]);
            if (node.Attributes[nameof(startDate)] != null)
                ctObj.startDate = XmlHelper.ReadDateTime(node.Attributes[nameof(startDate)]);
            if (node.Attributes[nameof(endDate)] != null)
                ctObj.endDate = XmlHelper.ReadDateTime(node.Attributes[nameof(endDate)]);
            if (node.Attributes[nameof(groupInterval)] != null)
                ctObj.groupInterval = XmlHelper.ReadDouble(node.Attributes[nameof(groupInterval)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(autoStart), this.autoStart);
            XmlHelper.WriteAttribute(sw, nameof(autoEnd), this.autoEnd);
            XmlHelper.WriteAttribute(sw, nameof(groupBy), this.groupBy.ToString());
            XmlHelper.WriteAttribute(sw, nameof(startNum), this.startNum);
            XmlHelper.WriteAttribute(sw, nameof(endNum), this.endNum);
            XmlHelper.WriteAttribute(sw, nameof(startDate), this.startDate);
            XmlHelper.WriteAttribute(sw, nameof(endDate), this.endDate);
            XmlHelper.WriteAttribute(sw, nameof(groupInterval), this.groupInterval);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_RangePr()
        {
            this.autoStart = true;
            this.autoEnd = true;
            this.groupBy = ST_GroupBy.range;
            this.groupInterval = 1D;
        }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool autoStart { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool autoEnd { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_GroupBy.range)]
        public ST_GroupBy groupBy { get; set; }

        [XmlAttribute()]
        public double startNum { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool startNumSpecified { get; set; }

        [XmlAttribute()]
        public double endNum { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool endNumSpecified { get; set; }

        [XmlAttribute()]
        public System.DateTime? startDate { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool startDateSpecified { get; set; }

        [XmlAttribute()]
        public System.DateTime? endDate { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool endDateSpecified { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(1D)]
        public double groupInterval { get; set; }
    }

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_GroupBy
    {

        /// <remarks/>
        range,

        /// <remarks/>
        seconds,

        /// <remarks/>
        minutes,

        /// <remarks/>
        hours,

        /// <remarks/>
        days,

        /// <remarks/>
        months,

        /// <remarks/>
        quarters,

        /// <remarks/>
        years,
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_DiscretePr
    {
        public static CT_DiscretePr Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_DiscretePr ctObj = new CT_DiscretePr();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.x = new List<CT_Index>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(x))
                    ctObj.x.Add(CT_Index.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.x != null)
            {
                foreach (CT_Index x in this.x)
                {
                    x.Write(sw, nameof(x));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_DiscretePr()
        {
            this.x = new List<CT_Index>();
        }

        [XmlElement(nameof(x), Order = 0)]
        public List<CT_Index> x { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_GroupItems
    {
        public static CT_GroupItems Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_GroupItems ctObj = new CT_GroupItems();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "e")
                {
                    ctObj.Items.Add(CT_Error.Parse(childNode, namespaceManager));
                    //ctObj.ItemsElementName.Add(ItemsChoiceType1.e);
                }
                else if (childNode.LocalName == "b")
                {
                    ctObj.Items.Add(CT_Boolean.Parse(childNode, namespaceManager));
                    //ctObj.ItemsElementName.Add(ItemsChoiceType1.b);
                }
                else if (childNode.LocalName == "d")
                {
                    ctObj.Items.Add(CT_DateTime.Parse(childNode, namespaceManager));
                    //ctObj.ItemsElementName.Add(ItemsChoiceType1.d);
                }
                else if (childNode.LocalName == "n")
                {
                    ctObj.Items.Add(CT_Number.Parse(childNode, namespaceManager));
                    //ctObj.ItemsElementName.Add(ItemsChoiceType1.n);
                }
                else if (childNode.LocalName == "m")
                {
                    ctObj.Items.Add(CT_Missing.Parse(childNode, namespaceManager));
                    //ctObj.ItemsElementName.Add(ItemsChoiceType1.m);
                }
                else if (childNode.LocalName == "s")
                {
                    ctObj.Items.Add(CT_String.Parse(childNode, namespaceManager));
                    //ctObj.ItemsElementName.Add(ItemsChoiceType1.s);
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            foreach (object o in this.Items)
            {
                if (o is CT_Error)
                    ((CT_Error)o).Write(sw, "e");
                else if (o is CT_Boolean)
                    ((CT_Boolean)o).Write(sw, "b");
                else if (o is CT_DateTime)
                    ((CT_DateTime)o).Write(sw, "d");
                else if (o is CT_Number)
                    ((CT_Number)o).Write(sw, "n");
                else if (o is CT_Missing)
                    ((CT_Missing)o).Write(sw, "m");
                else if (o is CT_String)
                    ((CT_String)o).Write(sw, "s");
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_GroupItems()
        {
            this.Items = new List<object>();
        }

        [XmlElement("b", typeof(CT_Boolean), Order = 0)]
        [XmlElement("d", typeof(CT_DateTime), Order = 0)]
        [XmlElement("e", typeof(CT_Error), Order = 0)]
        [XmlElement("m", typeof(CT_Missing), Order = 0)]
        [XmlElement("n", typeof(CT_Number), Order = 0)]
        [XmlElement("s", typeof(CT_String), Order = 0)]
        public List<object> Items { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_FieldsUsage
    {
        public static CT_FieldsUsage Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FieldsUsage ctObj = new CT_FieldsUsage();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.fieldUsage = new List<CT_FieldUsage>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(fieldUsage))
                    ctObj.fieldUsage.Add(CT_FieldUsage.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.fieldUsage != null)
            {
                foreach (CT_FieldUsage x in this.fieldUsage)
                {
                    x.Write(sw, nameof(fieldUsage));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_FieldsUsage()
        {
            this.fieldUsage = new List<CT_FieldUsage>();
        }

        [XmlElement(nameof(fieldUsage), Order = 0)]
        public List<CT_FieldUsage> fieldUsage { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_FieldUsage
    {
        public static CT_FieldUsage Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_FieldUsage ctObj = new CT_FieldUsage();
            if (node.Attributes[nameof(x)] != null)
                ctObj.x = XmlHelper.ReadInt(node.Attributes[nameof(x)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(x), this.x);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [XmlAttribute()]
        public int x { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_GroupLevels
    {
        public static CT_GroupLevels Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_GroupLevels ctObj = new CT_GroupLevels();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.groupLevel = new List<CT_GroupLevel>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(groupLevel))
                    ctObj.groupLevel.Add(CT_GroupLevel.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.groupLevel != null)
            {
                foreach (CT_GroupLevel x in this.groupLevel)
                {
                    x.Write(sw, nameof(groupLevel));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_GroupLevels()
        {
            this.groupLevel = new List<CT_GroupLevel>();
        }

        [XmlElement(nameof(groupLevel), Order = 0)]
        public List<CT_GroupLevel> groupLevel { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_GroupLevel
    {
        public static CT_GroupLevel Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_GroupLevel ctObj = new CT_GroupLevel();
            ctObj.uniqueName = XmlHelper.ReadString(node.Attributes[nameof(uniqueName)]);
            ctObj.caption = XmlHelper.ReadString(node.Attributes[nameof(caption)]);
            if (node.Attributes[nameof(user)] != null)
                ctObj.user = XmlHelper.ReadBool(node.Attributes[nameof(user)]);
            if (node.Attributes[nameof(customRollUp)] != null)
                ctObj.customRollUp = XmlHelper.ReadBool(node.Attributes[nameof(customRollUp)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(groups))
                    ctObj.groups = CT_Groups.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(uniqueName), this.uniqueName);
            XmlHelper.WriteAttribute(sw, nameof(caption), this.caption);
            XmlHelper.WriteAttribute(sw, nameof(user), this.user);
            XmlHelper.WriteAttribute(sw, nameof(customRollUp), this.customRollUp);
            sw.Write(">");
            if (this.groups != null)
                this.groups.Write(sw, nameof(groups));
            if (this.extLst != null)
                this.extLst.Write(sw, nameof(extLst));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_GroupLevel()
        {
            this.extLst = new CT_ExtensionList();
            this.groups = new CT_Groups();
            this.user = false;
            this.customRollUp = false;
        }

        [XmlElement(Order = 0)]
        public CT_Groups groups { get; set; }

        [XmlElement(Order = 1)]
        public CT_ExtensionList extLst { get; set; }

        [XmlAttribute()]
        public string uniqueName { get; set; }

        [XmlAttribute()]
        public string caption { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool user { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool customRollUp { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Groups
    {
        public static CT_Groups Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Groups ctObj = new CT_Groups();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.group = new List<CT_LevelGroup>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(group))
                    ctObj.group.Add(CT_LevelGroup.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.group != null)
            {
                foreach (CT_LevelGroup x in this.group)
                {
                    x.Write(sw, nameof(group));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Groups()
        {
            this.group = new List<CT_LevelGroup>();
        }

        [XmlElement(nameof(group), Order = 0)]
        public List<CT_LevelGroup> group { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_LevelGroup
    {
        public static CT_LevelGroup Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_LevelGroup ctObj = new CT_LevelGroup();
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            ctObj.uniqueName = XmlHelper.ReadString(node.Attributes[nameof(uniqueName)]);
            ctObj.caption = XmlHelper.ReadString(node.Attributes[nameof(caption)]);
            ctObj.uniqueParent = XmlHelper.ReadString(node.Attributes[nameof(uniqueParent)]);
            if (node.Attributes[nameof(id)] != null)
                ctObj.id = XmlHelper.ReadInt(node.Attributes[nameof(id)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(groupMembers))
                    ctObj.groupMembers = CT_GroupMembers.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(uniqueName), this.uniqueName);
            XmlHelper.WriteAttribute(sw, nameof(caption), this.caption);
            XmlHelper.WriteAttribute(sw, nameof(uniqueParent), this.uniqueParent);
            XmlHelper.WriteAttribute(sw, nameof(id), this.id);
            sw.Write(">");
            if (this.groupMembers != null)
                this.groupMembers.Write(sw, nameof(groupMembers));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_LevelGroup()
        {
            this.groupMembers = new CT_GroupMembers();
        }

        [XmlElement(Order = 0)]
        public CT_GroupMembers groupMembers { get; set; }

        [XmlAttribute()]
        public string name { get; set; }

        [XmlAttribute()]
        public string uniqueName { get; set; }

        [XmlAttribute()]
        public string caption { get; set; }

        [XmlAttribute()]
        public string uniqueParent { get; set; }

        [XmlAttribute()]
        public int id { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool idSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_GroupMembers
    {
        public static CT_GroupMembers Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_GroupMembers ctObj = new CT_GroupMembers();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.groupMember = new List<CT_GroupMember>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(groupMember))
                    ctObj.groupMember.Add(CT_GroupMember.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.groupMember != null)
            {
                foreach (CT_GroupMember x in this.groupMember)
                {
                    x.Write(sw, nameof(groupMember));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_GroupMembers()
        {
            this.groupMember = new List<CT_GroupMember>();
        }

        [XmlElement(nameof(groupMember), Order = 0)]
        public List<CT_GroupMember> groupMember { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_GroupMember
    {
        public static CT_GroupMember Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_GroupMember ctObj = new CT_GroupMember();
            ctObj.uniqueName = XmlHelper.ReadString(node.Attributes[nameof(uniqueName)]);
            if (node.Attributes[nameof(group)] != null)
                ctObj.group = XmlHelper.ReadBool(node.Attributes[nameof(group)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(uniqueName), this.uniqueName);
            XmlHelper.WriteAttribute(sw, nameof(group), this.group);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_GroupMember()
        {
            this.group = false;
        }

        [XmlAttribute()]
        public string uniqueName { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool group { get; set; }
    }
    
    
    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_CacheHierarchy
    {
        public static CT_CacheHierarchy Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CacheHierarchy ctObj = new CT_CacheHierarchy();
            ctObj.uniqueName = XmlHelper.ReadString(node.Attributes[nameof(uniqueName)]);
            ctObj.caption = XmlHelper.ReadString(node.Attributes[nameof(caption)]);
            if (node.Attributes[nameof(measure)] != null)
                ctObj.measure = XmlHelper.ReadBool(node.Attributes[nameof(measure)]);
            if (node.Attributes[nameof(set)] != null)
                ctObj.set = XmlHelper.ReadBool(node.Attributes[nameof(set)]);
            if (node.Attributes[nameof(parentSet)] != null)
                ctObj.parentSet = XmlHelper.ReadUInt(node.Attributes[nameof(parentSet)]);
            if (node.Attributes[nameof(iconSet)] != null)
                ctObj.iconSet = XmlHelper.ReadInt(node.Attributes[nameof(iconSet)]);
            if (node.Attributes[nameof(attribute)] != null)
                ctObj.attribute = XmlHelper.ReadBool(node.Attributes[nameof(attribute)]);
            if (node.Attributes[nameof(time)] != null)
                ctObj.time = XmlHelper.ReadBool(node.Attributes[nameof(time)]);
            if (node.Attributes[nameof(keyAttribute)] != null)
                ctObj.keyAttribute = XmlHelper.ReadBool(node.Attributes[nameof(keyAttribute)]);
            ctObj.defaultMemberUniqueName = XmlHelper.ReadString(node.Attributes[nameof(defaultMemberUniqueName)]);
            ctObj.allUniqueName = XmlHelper.ReadString(node.Attributes[nameof(allUniqueName)]);
            ctObj.allCaption = XmlHelper.ReadString(node.Attributes[nameof(allCaption)]);
            ctObj.dimensionUniqueName = XmlHelper.ReadString(node.Attributes[nameof(dimensionUniqueName)]);
            ctObj.displayFolder = XmlHelper.ReadString(node.Attributes[nameof(displayFolder)]);
            ctObj.measureGroup = XmlHelper.ReadString(node.Attributes[nameof(measureGroup)]);
            if (node.Attributes[nameof(measures)] != null)
                ctObj.measures = XmlHelper.ReadBool(node.Attributes[nameof(measures)]);
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            if (node.Attributes[nameof(oneField)] != null)
                ctObj.oneField = XmlHelper.ReadBool(node.Attributes[nameof(oneField)]);
            if (node.Attributes[nameof(memberValueDatatype)] != null)
                ctObj.memberValueDatatype = XmlHelper.ReadUShort(node.Attributes[nameof(memberValueDatatype)]);
            if (node.Attributes[nameof(unbalanced)] != null)
                ctObj.unbalanced = XmlHelper.ReadBool(node.Attributes[nameof(unbalanced)]);
            if (node.Attributes[nameof(unbalancedGroup)] != null)
                ctObj.unbalancedGroup = XmlHelper.ReadBool(node.Attributes[nameof(unbalancedGroup)]);
            if (node.Attributes[nameof(hidden)] != null)
                ctObj.hidden = XmlHelper.ReadBool(node.Attributes[nameof(hidden)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(fieldsUsage))
                    ctObj.fieldsUsage = CT_FieldsUsage.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(groupLevels))
                    ctObj.groupLevels = CT_GroupLevels.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(uniqueName), this.uniqueName);
            XmlHelper.WriteAttribute(sw, nameof(caption), this.caption);
            XmlHelper.WriteAttribute(sw, nameof(measure), this.measure);
            XmlHelper.WriteAttribute(sw, nameof(set), this.set);
            XmlHelper.WriteAttribute(sw, nameof(parentSet), this.parentSet);
            XmlHelper.WriteAttribute(sw, nameof(iconSet), this.iconSet);
            XmlHelper.WriteAttribute(sw, nameof(attribute), this.attribute);
            XmlHelper.WriteAttribute(sw, nameof(time), this.time);
            XmlHelper.WriteAttribute(sw, nameof(keyAttribute), this.keyAttribute);
            XmlHelper.WriteAttribute(sw, nameof(defaultMemberUniqueName), this.defaultMemberUniqueName);
            XmlHelper.WriteAttribute(sw, nameof(allUniqueName), this.allUniqueName);
            XmlHelper.WriteAttribute(sw, nameof(allCaption), this.allCaption);
            XmlHelper.WriteAttribute(sw, nameof(dimensionUniqueName), this.dimensionUniqueName);
            XmlHelper.WriteAttribute(sw, nameof(displayFolder), this.displayFolder);
            XmlHelper.WriteAttribute(sw, nameof(measureGroup), this.measureGroup);
            XmlHelper.WriteAttribute(sw, nameof(measures), this.measures);
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            XmlHelper.WriteAttribute(sw, nameof(oneField), this.oneField);
            XmlHelper.WriteAttribute(sw, nameof(memberValueDatatype), this.memberValueDatatype);
            XmlHelper.WriteAttribute(sw, nameof(unbalanced), this.unbalanced);
            XmlHelper.WriteAttribute(sw, nameof(unbalancedGroup), this.unbalancedGroup);
            XmlHelper.WriteAttribute(sw, nameof(hidden), this.hidden);
            sw.Write(">");
            if (this.fieldsUsage != null)
                this.fieldsUsage.Write(sw, nameof(fieldsUsage));
            if (this.groupLevels != null)
                this.groupLevels.Write(sw, nameof(groupLevels));
            if (this.extLst != null)
                this.extLst.Write(sw, nameof(extLst));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_CacheHierarchy()
        {
            this.extLst = new CT_ExtensionList();
            this.groupLevels = new CT_GroupLevels();
            this.fieldsUsage = new CT_FieldsUsage();
            this.measure = false;
            this.set = false;
            this.iconSet = 0;
            this.attribute = false;
            this.time = false;
            this.keyAttribute = false;
            this.measures = false;
            this.oneField = false;
            this.hidden = false;
        }

        [XmlElement(Order = 0)]
        public CT_FieldsUsage fieldsUsage { get; set; }

        [XmlElement(Order = 1)]
        public CT_GroupLevels groupLevels { get; set; }

        [XmlElement(Order = 2)]
        public CT_ExtensionList extLst { get; set; }

        [XmlAttribute()]
        public string uniqueName { get; set; }

        [XmlAttribute()]
        public string caption { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool measure { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool set { get; set; }

        [XmlAttribute()]
        public uint parentSet { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool parentSetSpecified { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public int iconSet { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool attribute { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool time { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool keyAttribute { get; set; }

        [XmlAttribute()]
        public string defaultMemberUniqueName { get; set; }

        [XmlAttribute()]
        public string allUniqueName { get; set; }

        [XmlAttribute()]
        public string allCaption { get; set; }

        [XmlAttribute()]
        public string dimensionUniqueName { get; set; }

        [XmlAttribute()]
        public string displayFolder { get; set; }

        [XmlAttribute()]
        public string measureGroup { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool measures { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool oneField { get; set; }

        [XmlAttribute()]
        public ushort memberValueDatatype { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool memberValueDatatypeSpecified { get; set; }

        [XmlAttribute()]
        public bool unbalanced { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool unbalancedSpecified { get; set; }

        [XmlAttribute()]
        public bool unbalancedGroup { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool unbalancedGroupSpecified { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool hidden { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PCDKPIs
    {
        public static CT_PCDKPIs Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PCDKPIs ctObj = new CT_PCDKPIs();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.kpi = new List<CT_PCDKPI>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(kpi))
                    ctObj.kpi.Add(CT_PCDKPI.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.kpi != null)
            {
                foreach (CT_PCDKPI x in this.kpi)
                {
                    x.Write(sw, nameof(kpi));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_PCDKPIs()
        {
            this.kpi = new List<CT_PCDKPI>();
        }

        [XmlElement(nameof(kpi), Order = 0)]
        public List<CT_PCDKPI> kpi { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PCDKPI
    {
        public static CT_PCDKPI Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PCDKPI ctObj = new CT_PCDKPI();
            ctObj.uniqueName = XmlHelper.ReadString(node.Attributes[nameof(uniqueName)]);
            ctObj.caption = XmlHelper.ReadString(node.Attributes[nameof(caption)]);
            ctObj.displayFolder = XmlHelper.ReadString(node.Attributes[nameof(displayFolder)]);
            ctObj.measureGroup = XmlHelper.ReadString(node.Attributes[nameof(measureGroup)]);
            ctObj.parent = XmlHelper.ReadString(node.Attributes[nameof(parent)]);
            ctObj.value = XmlHelper.ReadString(node.Attributes[nameof(value)]);
            ctObj.goal = XmlHelper.ReadString(node.Attributes[nameof(goal)]);
            ctObj.status = XmlHelper.ReadString(node.Attributes[nameof(status)]);
            ctObj.trend = XmlHelper.ReadString(node.Attributes[nameof(trend)]);
            ctObj.weight = XmlHelper.ReadString(node.Attributes[nameof(weight)]);
            ctObj.time = XmlHelper.ReadString(node.Attributes[nameof(time)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(uniqueName), this.uniqueName);
            XmlHelper.WriteAttribute(sw, nameof(caption), this.caption);
            XmlHelper.WriteAttribute(sw, nameof(displayFolder), this.displayFolder);
            XmlHelper.WriteAttribute(sw, nameof(measureGroup), this.measureGroup);
            XmlHelper.WriteAttribute(sw, nameof(parent), this.parent);
            XmlHelper.WriteAttribute(sw, nameof(value), this.value);
            XmlHelper.WriteAttribute(sw, nameof(goal), this.goal);
            XmlHelper.WriteAttribute(sw, nameof(status), this.status);
            XmlHelper.WriteAttribute(sw, nameof(trend), this.trend);
            XmlHelper.WriteAttribute(sw, nameof(weight), this.weight);
            XmlHelper.WriteAttribute(sw, nameof(time), this.time);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [XmlAttribute()]
        public string uniqueName { get; set; }

        [XmlAttribute()]
        public string caption { get; set; }

        [XmlAttribute()]
        public string displayFolder { get; set; }

        [XmlAttribute()]
        public string measureGroup { get; set; }

        [XmlAttribute()]
        public string parent { get; set; }

        [XmlAttribute()]
        public string value { get; set; }

        [XmlAttribute()]
        public string goal { get; set; }

        [XmlAttribute()]
        public string status { get; set; }

        [XmlAttribute()]
        public string trend { get; set; }

        [XmlAttribute()]
        public string weight { get; set; }

        [XmlAttribute()]
        public string time { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_TupleCache
    {
        public static CT_TupleCache Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_TupleCache ctObj = new CT_TupleCache();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(entries))
                    ctObj.entries = CT_PCDSDTCEntries.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(sets))
                    ctObj.sets = CT_Sets.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(queryCache))
                    ctObj.queryCache = CT_QueryCache.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(serverFormats))
                    ctObj.serverFormats = CT_ServerFormats.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            if (this.entries != null)
                this.entries.Write(sw, nameof(entries));
            if (this.sets != null)
                this.sets.Write(sw, nameof(sets));
            if (this.queryCache != null)
                this.queryCache.Write(sw, nameof(queryCache));
            if (this.serverFormats != null)
                this.serverFormats.Write(sw, nameof(serverFormats));
            if (this.extLst != null)
                this.extLst.Write(sw, nameof(extLst));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_TupleCache()
        {
            this.extLst = new CT_ExtensionList();
            this.serverFormats = new CT_ServerFormats();
            this.queryCache = new CT_QueryCache();
            this.sets = new CT_Sets();
            this.entries = new CT_PCDSDTCEntries();
        }

        [XmlElement(Order = 0)]
        public CT_PCDSDTCEntries entries { get; set; }

        [XmlElement(Order = 1)]
        public CT_Sets sets { get; set; }

        [XmlElement(Order = 2)]
        public CT_QueryCache queryCache { get; set; }

        [XmlElement(Order = 3)]
        public CT_ServerFormats serverFormats { get; set; }

        [XmlElement(Order = 4)]
        public CT_ExtensionList extLst { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PCDSDTCEntries
    {
        public static CT_PCDSDTCEntries Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PCDSDTCEntries ctObj = new CT_PCDSDTCEntries();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == "m")
                {
                    ctObj.Items.Add(CT_Missing.Parse(childNode, namespaceManager));
                    //ctObj.ItemsElementName.Add(ItemsChoiceType1.m);
                }
                else if (childNode.LocalName == "n")
                {
                    ctObj.Items.Add(CT_Number.Parse(childNode, namespaceManager));
                    //ctObj.ItemsElementName.Add(ItemsChoiceType1.n);
                }
                else if (childNode.LocalName == "e")
                {
                    ctObj.Items.Add(CT_Error.Parse(childNode, namespaceManager));
                    //ctObj.ItemsElementName.Add(ItemsChoiceType1.e);
                }
                else if (childNode.LocalName == "s")
                {
                    ctObj.Items.Add(CT_String.Parse(childNode, namespaceManager));
                    //ctObj.ItemsElementName.Add(ItemsChoiceType1.s);
                }
            }
            return ctObj;
        }

        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            foreach (object o in this.Items)
            {
                if (o is CT_Missing)
                    ((CT_Missing)o).Write(sw, "m");
                else if (o is CT_Number)
                    ((CT_Number)o).Write(sw, "n");
                else if (o is CT_Error)
                    ((CT_Error)o).Write(sw, "e");
                else if (o is CT_String)
                    ((CT_String)o).Write(sw, "s");
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_PCDSDTCEntries()
        {
            this.Items = new List<object>();
        }

        [XmlElement("e", typeof(CT_Error), Order = 0)]
        [XmlElement("m", typeof(CT_Missing), Order = 0)]
        [XmlElement("n", typeof(CT_Number), Order = 0)]
        [XmlElement("s", typeof(CT_String), Order = 0)]
        public List<object> Items { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Sets
    {
        public static CT_Sets Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Sets ctObj = new CT_Sets();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.set = new List<CT_Set>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(set))
                    ctObj.set.Add(CT_Set.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.set != null)
            {
                foreach (CT_Set x in this.set)
                {
                    x.Write(sw, nameof(set));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Sets()
        {
            this.set = new List<CT_Set>();
        }

        [XmlElement(nameof(set), Order = 0)]
        public List<CT_Set> set { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Set
    {
        public static CT_Set Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Set ctObj = new CT_Set();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            if (node.Attributes[nameof(maxRank)] != null)
                ctObj.maxRank = XmlHelper.ReadInt(node.Attributes[nameof(maxRank)]);
            ctObj.setDefinition = XmlHelper.ReadString(node.Attributes[nameof(setDefinition)]);
            if (node.Attributes[nameof(sortType)] != null)
                ctObj.sortType = (ST_SortType)Enum.Parse(typeof(ST_SortType), node.Attributes[nameof(sortType)].Value);
            if (node.Attributes[nameof(queryFailed)] != null)
                ctObj.queryFailed = XmlHelper.ReadBool(node.Attributes[nameof(queryFailed)]);
            ctObj.tpls = new List<CT_Tuples>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(sortByTuple))
                    ctObj.sortByTuple = CT_Tuples.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(tpls))
                    ctObj.tpls.Add(CT_Tuples.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            XmlHelper.WriteAttribute(sw, nameof(maxRank), this.maxRank);
            XmlHelper.WriteAttribute(sw, nameof(setDefinition), this.setDefinition);
            XmlHelper.WriteAttribute(sw, nameof(sortType), this.sortType.ToString());
            XmlHelper.WriteAttribute(sw, nameof(queryFailed), this.queryFailed);
            sw.Write(">");
            if (this.sortByTuple != null)
                this.sortByTuple.Write(sw, nameof(sortByTuple));
            if (this.tpls != null)
            {
                foreach (CT_Tuples x in this.tpls)
                {
                    x.Write(sw, nameof(tpls));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Set()
        {
            this.sortByTuple = new CT_Tuples();
            this.tpls = new List<CT_Tuples>();
            this.sortType = ST_SortType.none;
            this.queryFailed = false;
        }

        [XmlElement(nameof(tpls), Order = 0)]
        public List<CT_Tuples> tpls { get; set; }

        [XmlElement(Order = 1)]
        public CT_Tuples sortByTuple { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }

        [XmlAttribute()]
        public int maxRank { get; set; }

        [XmlAttribute()]
        public string setDefinition { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_SortType.none)]
        public ST_SortType sortType { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool queryFailed { get; set; }
    }

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_SortType
    {

        /// <remarks/>
        none,

        /// <remarks/>
        ascending,

        /// <remarks/>
        descending,

        /// <remarks/>
        ascendingAlpha,

        /// <remarks/>
        descendingAlpha,

        /// <remarks/>
        ascendingNatural,

        /// <remarks/>
        descendingNatural,
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_QueryCache
    {
        public static CT_QueryCache Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_QueryCache ctObj = new CT_QueryCache();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.query = new List<CT_Query>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(query))
                    ctObj.query.Add(CT_Query.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.query != null)
            {
                foreach (CT_Query x in this.query)
                {
                    x.Write(sw, nameof(query));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_QueryCache()
        {
            this.query = new List<CT_Query>();
        }

        [XmlElement(nameof(query), Order = 0)]
        public List<CT_Query> query { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Query
    {
        public static CT_Query Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Query ctObj = new CT_Query();
            ctObj.mdx = XmlHelper.ReadString(node.Attributes[nameof(mdx)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(tpls))
                    ctObj.tpls = CT_Tuples.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(mdx), this.mdx);
            sw.Write(">");
            if (this.tpls != null)
                this.tpls.Write(sw, nameof(tpls));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Query()
        {
            this.tpls = new CT_Tuples();
        }

        [XmlElement(Order = 0)]
        public CT_Tuples tpls { get; set; }

        [XmlAttribute()]
        public string mdx { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_ServerFormats
    {
        public static CT_ServerFormats Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ServerFormats ctObj = new CT_ServerFormats();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.serverFormat = new List<CT_ServerFormat>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(serverFormat))
                    ctObj.serverFormat.Add(CT_ServerFormat.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.serverFormat != null)
            {
                foreach (CT_ServerFormat x in this.serverFormat)
                {
                    x.Write(sw, nameof(serverFormat));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_ServerFormats()
        {
            this.serverFormat = new List<CT_ServerFormat>();
        }

        [XmlElement(nameof(serverFormat), Order = 0)]
        public List<CT_ServerFormat> serverFormat { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_ServerFormat
    {
        public static CT_ServerFormat Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ServerFormat ctObj = new CT_ServerFormat();
            ctObj.culture = XmlHelper.ReadString(node.Attributes[nameof(culture)]);
            ctObj.format = XmlHelper.ReadString(node.Attributes[nameof(format)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(culture), this.culture);
            XmlHelper.WriteAttribute(sw, nameof(format), this.format);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [XmlAttribute()]
        public string culture { get; set; }

        [XmlAttribute()]
        public string format { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_CalculatedItems
    {
        public static CT_CalculatedItems Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CalculatedItems ctObj = new CT_CalculatedItems();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.calculatedItem = new List<CT_CalculatedItem>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(calculatedItem))
                    ctObj.calculatedItem.Add(CT_CalculatedItem.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.calculatedItem != null)
            {
                foreach (CT_CalculatedItem x in this.calculatedItem)
                {
                    x.Write(sw, nameof(calculatedItem));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_CalculatedItems()
        {
            this.calculatedItem = new List<CT_CalculatedItem>();
        }

        [XmlElement(nameof(calculatedItem), Order = 0)]
        public List<CT_CalculatedItem> calculatedItem { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_CalculatedItem
    {
        public static CT_CalculatedItem Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CalculatedItem ctObj = new CT_CalculatedItem();
            if (node.Attributes[nameof(field)] != null)
                ctObj.field = XmlHelper.ReadUInt(node.Attributes[nameof(field)]);
            ctObj.formula = XmlHelper.ReadString(node.Attributes[nameof(formula)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(pivotArea))
                    ctObj.pivotArea = CT_PivotArea.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(field), this.field);
            XmlHelper.WriteAttribute(sw, nameof(formula), this.formula);
            sw.Write(">");
            if (this.pivotArea != null)
                this.pivotArea.Write(sw, nameof(pivotArea));
            if (this.extLst != null)
                this.extLst.Write(sw, nameof(extLst));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_CalculatedItem()
        {
            this.extLst = new CT_ExtensionList();
            this.pivotArea = new CT_PivotArea();
        }

        [XmlElement(Order = 0)]
        public CT_PivotArea pivotArea { get; set; }

        [XmlElement(Order = 1)]
        public CT_ExtensionList extLst { get; set; }

        [XmlAttribute()]
        public uint field { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool fieldSpecified { get; set; }

        [XmlAttribute()]
        public string formula { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_CalculatedMembers
    {
        public static CT_CalculatedMembers Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CalculatedMembers ctObj = new CT_CalculatedMembers();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.calculatedMember = new List<CT_CalculatedMember>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(calculatedMember))
                    ctObj.calculatedMember.Add(CT_CalculatedMember.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.calculatedMember != null)
            {
                foreach (CT_CalculatedMember x in this.calculatedMember)
                {
                    x.Write(sw, nameof(calculatedMember));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_CalculatedMembers()
        {
            this.calculatedMember = new List<CT_CalculatedMember>();
        }

        [XmlElement(nameof(calculatedMember), Order = 0)]
        public List<CT_CalculatedMember> calculatedMember { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_CalculatedMember
    {
        public static CT_CalculatedMember Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_CalculatedMember ctObj = new CT_CalculatedMember();
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            ctObj.mdx = XmlHelper.ReadString(node.Attributes[nameof(mdx)]);
            ctObj.memberName = XmlHelper.ReadString(node.Attributes[nameof(memberName)]);
            ctObj.hierarchy = XmlHelper.ReadString(node.Attributes[nameof(hierarchy)]);
            ctObj.parent = XmlHelper.ReadString(node.Attributes[nameof(parent)]);
            if (node.Attributes[nameof(solveOrder)] != null)
                ctObj.solveOrder = XmlHelper.ReadInt(node.Attributes[nameof(solveOrder)]);
            if (node.Attributes[nameof(set)] != null)
                ctObj.set = XmlHelper.ReadBool(node.Attributes[nameof(set)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(mdx), this.mdx);
            XmlHelper.WriteAttribute(sw, nameof(memberName), this.memberName);
            XmlHelper.WriteAttribute(sw, nameof(hierarchy), this.hierarchy);
            XmlHelper.WriteAttribute(sw, nameof(parent), this.parent);
            XmlHelper.WriteAttribute(sw, nameof(solveOrder), this.solveOrder);
            XmlHelper.WriteAttribute(sw, nameof(set), this.set);
            sw.Write(">");
            if (this.extLst != null)
                this.extLst.Write(sw, nameof(extLst));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_CalculatedMember()
        {
            this.extLst = new CT_ExtensionList();
            this.solveOrder = 0;
            this.set = false;
        }

        [XmlElement(Order = 0)]
        public CT_ExtensionList extLst { get; set; }

        [XmlAttribute()]
        public string name { get; set; }

        [XmlAttribute()]
        public string mdx { get; set; }

        [XmlAttribute()]
        public string memberName { get; set; }

        [XmlAttribute()]
        public string hierarchy { get; set; }

        [XmlAttribute()]
        public string parent { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public int solveOrder { get; set; }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool set { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Dimensions
    {
        public static CT_Dimensions Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Dimensions ctObj = new CT_Dimensions();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.dimension = new List<CT_PivotDimension>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(dimension))
                    ctObj.dimension.Add(CT_PivotDimension.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.dimension != null)
            {
                foreach (CT_PivotDimension x in this.dimension)
                {
                    x.Write(sw, nameof(dimension));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Dimensions()
        {
            this.dimension = new List<CT_PivotDimension>();
        }

        [XmlElement(nameof(dimension), Order = 0)]
        public List<CT_PivotDimension> dimension { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotDimension
    {
        public static CT_PivotDimension Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PivotDimension ctObj = new CT_PivotDimension();
            if (node.Attributes[nameof(measure)] != null)
                ctObj.measure = XmlHelper.ReadBool(node.Attributes[nameof(measure)]);
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            ctObj.uniqueName = XmlHelper.ReadString(node.Attributes[nameof(uniqueName)]);
            ctObj.caption = XmlHelper.ReadString(node.Attributes[nameof(caption)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(measure), this.measure);
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(uniqueName), this.uniqueName);
            XmlHelper.WriteAttribute(sw, nameof(caption), this.caption);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_PivotDimension()
        {
            this.measure = false;
        }

        [XmlAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool measure { get; set; }

        [XmlAttribute()]
        public string name { get; set; }

        [XmlAttribute()]
        public string uniqueName { get; set; }

        [XmlAttribute()]
        public string caption { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_MeasureGroups
    {
        public static CT_MeasureGroups Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MeasureGroups ctObj = new CT_MeasureGroups();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.measureGroup = new List<CT_MeasureGroup>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(measureGroup))
                    ctObj.measureGroup.Add(CT_MeasureGroup.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.measureGroup != null)
            {
                foreach (CT_MeasureGroup x in this.measureGroup)
                {
                    x.Write(sw, nameof(measureGroup));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_MeasureGroups()
        {
            this.measureGroup = new List<CT_MeasureGroup>();
        }

        [XmlElement(nameof(measureGroup), Order = 0)]
        public List<CT_MeasureGroup> measureGroup { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_MeasureGroup
    {
        public static CT_MeasureGroup Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MeasureGroup ctObj = new CT_MeasureGroup();
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            ctObj.caption = XmlHelper.ReadString(node.Attributes[nameof(caption)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(caption), this.caption);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [XmlAttribute()]
        public string name { get; set; }

        [XmlAttribute()]
        public string caption { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_MeasureDimensionMaps
    {
        public static CT_MeasureDimensionMaps Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MeasureDimensionMaps ctObj = new CT_MeasureDimensionMaps();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.map = new List<CT_MeasureDimensionMap>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(map))
                    ctObj.map.Add(CT_MeasureDimensionMap.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.map != null)
            {
                foreach (CT_MeasureDimensionMap x in this.map)
                {
                    x.Write(sw, nameof(map));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_MeasureDimensionMaps()
        {
            this.map = new List<CT_MeasureDimensionMap>();
        }

        [XmlElement(nameof(map), Order = 0)]
        public List<CT_MeasureDimensionMap> map { get; set; }

        [XmlAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_MeasureDimensionMap
    {
        public static CT_MeasureDimensionMap Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MeasureDimensionMap ctObj = new CT_MeasureDimensionMap();
            if (node.Attributes[nameof(measureGroup)] != null)
                ctObj.measureGroup = XmlHelper.ReadUInt(node.Attributes[nameof(measureGroup)]);
            if (node.Attributes[nameof(dimension)] != null)
                ctObj.dimension = XmlHelper.ReadUInt(node.Attributes[nameof(dimension)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(measureGroup), this.measureGroup);
            XmlHelper.WriteAttribute(sw, nameof(dimension), this.dimension);
            sw.Write("/>");
        }

        [XmlAttribute()]
        public uint measureGroup { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool measureGroupSpecified { get; set; }

        [XmlAttribute()]
        public uint dimension { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool dimensionSpecified { get; set; }
    }
}
