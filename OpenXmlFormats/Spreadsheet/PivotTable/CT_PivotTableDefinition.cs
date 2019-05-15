using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;
using System.Xml;
using NPOI.OpenXml4Net.Util;
using System.IO;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot("pivotTableDefinition", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public partial class CT_PivotTableDefinition
    {
        public static CT_PivotTableDefinition Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PivotTableDefinition ctObj = new CT_PivotTableDefinition();
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            if (node.Attributes[nameof(cacheId)] != null)
                ctObj.cacheId = XmlHelper.ReadUInt(node.Attributes[nameof(cacheId)]);
            if (node.Attributes[nameof(dataOnRows)] != null)
                ctObj.dataOnRows = XmlHelper.ReadBool(node.Attributes[nameof(dataOnRows)]);
            if (node.Attributes[nameof(dataPosition)] != null)
                ctObj.dataPosition = XmlHelper.ReadUInt(node.Attributes[nameof(dataPosition)]);
            if (node.Attributes[nameof(autoFormatId)] != null)
                ctObj.autoFormatId = XmlHelper.ReadUInt(node.Attributes[nameof(autoFormatId)]);
            if (node.Attributes[nameof(applyNumberFormats)] != null)
                ctObj.applyNumberFormats = XmlHelper.ReadBool(node.Attributes[nameof(applyNumberFormats)]);
            if (node.Attributes[nameof(applyBorderFormats)] != null)
                ctObj.applyBorderFormats = XmlHelper.ReadBool(node.Attributes[nameof(applyBorderFormats)]);
            if (node.Attributes[nameof(applyFontFormats)] != null)
                ctObj.applyFontFormats = XmlHelper.ReadBool(node.Attributes[nameof(applyFontFormats)]);
            if (node.Attributes[nameof(applyPatternFormats)] != null)
                ctObj.applyPatternFormats = XmlHelper.ReadBool(node.Attributes[nameof(applyPatternFormats)]);
            if (node.Attributes[nameof(applyAlignmentFormats)] != null)
                ctObj.applyAlignmentFormats = XmlHelper.ReadBool(node.Attributes[nameof(applyAlignmentFormats)]);
            if (node.Attributes[nameof(applyWidthHeightFormats)] != null)
                ctObj.applyWidthHeightFormats = XmlHelper.ReadBool(node.Attributes[nameof(applyWidthHeightFormats)]);
            ctObj.dataCaption = XmlHelper.ReadString(node.Attributes[nameof(dataCaption)]);
            ctObj.grandTotalCaption = XmlHelper.ReadString(node.Attributes[nameof(grandTotalCaption)]);
            ctObj.errorCaption = XmlHelper.ReadString(node.Attributes[nameof(errorCaption)]);
            if (node.Attributes[nameof(showError)] != null)
                ctObj.showError = XmlHelper.ReadBool(node.Attributes[nameof(showError)]);
            ctObj.missingCaption = XmlHelper.ReadString(node.Attributes[nameof(missingCaption)]);
            if (node.Attributes[nameof(showMissing)] != null)
                ctObj.showMissing = XmlHelper.ReadBool(node.Attributes[nameof(showMissing)]);
            ctObj.pageStyle = XmlHelper.ReadString(node.Attributes[nameof(pageStyle)]);
            ctObj.pivotTableStyle = XmlHelper.ReadString(node.Attributes[nameof(pivotTableStyle)]);
            ctObj.vacatedStyle = XmlHelper.ReadString(node.Attributes[nameof(vacatedStyle)]);
            ctObj.tag = XmlHelper.ReadString(node.Attributes[nameof(tag)]);
            if (node.Attributes[nameof(updatedVersion)] != null)
                ctObj.updatedVersion = XmlHelper.ReadByte(node.Attributes[nameof(updatedVersion)]);
            if (node.Attributes[nameof(minRefreshableVersion)] != null)
                ctObj.minRefreshableVersion = XmlHelper.ReadByte(node.Attributes[nameof(minRefreshableVersion)]);
            if (node.Attributes[nameof(asteriskTotals)] != null)
                ctObj.asteriskTotals = XmlHelper.ReadBool(node.Attributes[nameof(asteriskTotals)]);
            if (node.Attributes[nameof(showItems)] != null)
                ctObj.showItems = XmlHelper.ReadBool(node.Attributes[nameof(showItems)]);
            if (node.Attributes[nameof(editData)] != null)
                ctObj.editData = XmlHelper.ReadBool(node.Attributes[nameof(editData)]);
            if (node.Attributes[nameof(disableFieldList)] != null)
                ctObj.disableFieldList = XmlHelper.ReadBool(node.Attributes[nameof(disableFieldList)]);
            if (node.Attributes[nameof(showCalcMbrs)] != null)
                ctObj.showCalcMbrs = XmlHelper.ReadBool(node.Attributes[nameof(showCalcMbrs)]);
            if (node.Attributes[nameof(visualTotals)] != null)
                ctObj.visualTotals = XmlHelper.ReadBool(node.Attributes[nameof(visualTotals)]);
            if (node.Attributes[nameof(showMultipleLabel)] != null)
                ctObj.showMultipleLabel = XmlHelper.ReadBool(node.Attributes[nameof(showMultipleLabel)]);
            if (node.Attributes[nameof(showDataDropDown)] != null)
                ctObj.showDataDropDown = XmlHelper.ReadBool(node.Attributes[nameof(showDataDropDown)]);
            if (node.Attributes[nameof(showDrill)] != null)
                ctObj.showDrill = XmlHelper.ReadBool(node.Attributes[nameof(showDrill)]);
            if (node.Attributes[nameof(printDrill)] != null)
                ctObj.printDrill = XmlHelper.ReadBool(node.Attributes[nameof(printDrill)]);
            if (node.Attributes[nameof(showMemberPropertyTips)] != null)
                ctObj.showMemberPropertyTips = XmlHelper.ReadBool(node.Attributes[nameof(showMemberPropertyTips)]);
            if (node.Attributes[nameof(showDataTips)] != null)
                ctObj.showDataTips = XmlHelper.ReadBool(node.Attributes[nameof(showDataTips)]);
            if (node.Attributes[nameof(enableWizard)] != null)
                ctObj.enableWizard = XmlHelper.ReadBool(node.Attributes[nameof(enableWizard)]);
            if (node.Attributes[nameof(enableDrill)] != null)
                ctObj.enableDrill = XmlHelper.ReadBool(node.Attributes[nameof(enableDrill)]);
            if (node.Attributes[nameof(enableFieldProperties)] != null)
                ctObj.enableFieldProperties = XmlHelper.ReadBool(node.Attributes[nameof(enableFieldProperties)]);
            if (node.Attributes[nameof(preserveFormatting)] != null)
                ctObj.preserveFormatting = XmlHelper.ReadBool(node.Attributes[nameof(preserveFormatting)]);
            if (node.Attributes[nameof(useAutoFormatting)] != null)
                ctObj.useAutoFormatting = XmlHelper.ReadBool(node.Attributes[nameof(useAutoFormatting)]);
            if (node.Attributes[nameof(pageWrap)] != null)
                ctObj.pageWrap = XmlHelper.ReadUInt(node.Attributes[nameof(pageWrap)]);
            if (node.Attributes[nameof(pageOverThenDown)] != null)
                ctObj.pageOverThenDown = XmlHelper.ReadBool(node.Attributes[nameof(pageOverThenDown)]);
            if (node.Attributes[nameof(subtotalHiddenItems)] != null)
                ctObj.subtotalHiddenItems = XmlHelper.ReadBool(node.Attributes[nameof(subtotalHiddenItems)]);
            if (node.Attributes[nameof(rowGrandTotals)] != null)
                ctObj.rowGrandTotals = XmlHelper.ReadBool(node.Attributes[nameof(rowGrandTotals)]);
            if (node.Attributes[nameof(colGrandTotals)] != null)
                ctObj.colGrandTotals = XmlHelper.ReadBool(node.Attributes[nameof(colGrandTotals)]);
            if (node.Attributes[nameof(fieldPrintTitles)] != null)
                ctObj.fieldPrintTitles = XmlHelper.ReadBool(node.Attributes[nameof(fieldPrintTitles)]);
            if (node.Attributes[nameof(itemPrintTitles)] != null)
                ctObj.itemPrintTitles = XmlHelper.ReadBool(node.Attributes[nameof(itemPrintTitles)]);
            if (node.Attributes[nameof(mergeItem)] != null)
                ctObj.mergeItem = XmlHelper.ReadBool(node.Attributes[nameof(mergeItem)]);
            if (node.Attributes[nameof(showDropZones)] != null)
                ctObj.showDropZones = XmlHelper.ReadBool(node.Attributes[nameof(showDropZones)]);
            if (node.Attributes[nameof(createdVersion)] != null)
                ctObj.createdVersion = XmlHelper.ReadByte(node.Attributes[nameof(createdVersion)]);
            if (node.Attributes[nameof(indent)] != null)
                ctObj.indent = XmlHelper.ReadUInt(node.Attributes[nameof(indent)]);
            if (node.Attributes[nameof(showEmptyRow)] != null)
                ctObj.showEmptyRow = XmlHelper.ReadBool(node.Attributes[nameof(showEmptyRow)]);
            if (node.Attributes[nameof(showEmptyCol)] != null)
                ctObj.showEmptyCol = XmlHelper.ReadBool(node.Attributes[nameof(showEmptyCol)]);
            if (node.Attributes[nameof(showHeaders)] != null)
                ctObj.showHeaders = XmlHelper.ReadBool(node.Attributes[nameof(showHeaders)]);
            if (node.Attributes[nameof(compact)] != null)
                ctObj.compact = XmlHelper.ReadBool(node.Attributes[nameof(compact)]);
            if (node.Attributes[nameof(outline)] != null)
                ctObj.outline = XmlHelper.ReadBool(node.Attributes[nameof(outline)]);
            if (node.Attributes[nameof(outlineData)] != null)
                ctObj.outlineData = XmlHelper.ReadBool(node.Attributes[nameof(outlineData)]);
            if (node.Attributes[nameof(compactData)] != null)
                ctObj.compactData = XmlHelper.ReadBool(node.Attributes[nameof(compactData)]);
            if (node.Attributes[nameof(published)] != null)
                ctObj.published = XmlHelper.ReadBool(node.Attributes[nameof(published)]);
            if (node.Attributes[nameof(gridDropZones)] != null)
                ctObj.gridDropZones = XmlHelper.ReadBool(node.Attributes[nameof(gridDropZones)]);
            if (node.Attributes[nameof(immersive)] != null)
                ctObj.immersive = XmlHelper.ReadBool(node.Attributes[nameof(immersive)]);
            if (node.Attributes[nameof(multipleFieldFilters)] != null)
                ctObj.multipleFieldFilters = XmlHelper.ReadBool(node.Attributes[nameof(multipleFieldFilters)]);
            if (node.Attributes[nameof(chartFormat)] != null)
                ctObj.chartFormat = XmlHelper.ReadUInt(node.Attributes[nameof(chartFormat)]);
            ctObj.rowHeaderCaption = XmlHelper.ReadString(node.Attributes[nameof(rowHeaderCaption)]);
            ctObj.colHeaderCaption = XmlHelper.ReadString(node.Attributes[nameof(colHeaderCaption)]);
            if (node.Attributes[nameof(fieldListSortAscending)] != null)
                ctObj.fieldListSortAscending = XmlHelper.ReadBool(node.Attributes[nameof(fieldListSortAscending)]);
            if (node.Attributes[nameof(mdxSubqueries)] != null)
                ctObj.mdxSubqueries = XmlHelper.ReadBool(node.Attributes[nameof(mdxSubqueries)]);
            if (node.Attributes[nameof(customListSort)] != null)
                ctObj.customListSort = XmlHelper.ReadBool(node.Attributes[nameof(customListSort)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(location))
                    ctObj.location = CT_Location.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(pivotFields))
                    ctObj.pivotFields = CT_PivotFields.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(rowFields))
                    ctObj.rowFields = CT_RowFields.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(rowItems))
                    ctObj.rowItems = CT_rowItems.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(colFields))
                    ctObj.colFields = CT_ColFields.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(colItems))
                    ctObj.colItems = CT_colItems.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(pageFields))
                    ctObj.pageFields = CT_PageFields.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(dataFields))
                    ctObj.dataFields = CT_DataFields.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(formats))
                    ctObj.formats = CT_Formats.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(conditionalFormats))
                    ctObj.conditionalFormats = CT_ConditionalFormats.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(chartFormats))
                    ctObj.chartFormats = CT_ChartFormats.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(pivotHierarchies))
                    ctObj.pivotHierarchies = CT_PivotHierarchies.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(pivotTableStyleInfo))
                    ctObj.pivotTableStyleInfo = CT_PivotTableStyle.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(filters))
                    ctObj.filters = CT_PivotFilters.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(rowHierarchiesUsage))
                    ctObj.rowHierarchiesUsage = CT_RowHierarchiesUsage.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(colHierarchiesUsage))
                    ctObj.colHierarchiesUsage = CT_ColHierarchiesUsage.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw)
        {
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write("<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
            sw.Write("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ");
            sw.Write("xmlns:s=\"http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes\" ");
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(cacheId), this.cacheId);
            XmlHelper.WriteAttribute(sw, nameof(dataOnRows), this.dataOnRows);
            XmlHelper.WriteAttribute(sw, nameof(dataPosition), this.dataPosition);
            XmlHelper.WriteAttribute(sw, nameof(autoFormatId), this.autoFormatId);
            XmlHelper.WriteAttribute(sw, nameof(applyNumberFormats), this.applyNumberFormats);
            XmlHelper.WriteAttribute(sw, nameof(applyBorderFormats), this.applyBorderFormats);
            XmlHelper.WriteAttribute(sw, nameof(applyFontFormats), this.applyFontFormats);
            XmlHelper.WriteAttribute(sw, nameof(applyPatternFormats), this.applyPatternFormats);
            XmlHelper.WriteAttribute(sw, nameof(applyAlignmentFormats), this.applyAlignmentFormats);
            XmlHelper.WriteAttribute(sw, nameof(applyWidthHeightFormats), this.applyWidthHeightFormats);
            XmlHelper.WriteAttribute(sw, nameof(dataCaption), this.dataCaption);
            XmlHelper.WriteAttribute(sw, nameof(grandTotalCaption), this.grandTotalCaption);
            XmlHelper.WriteAttribute(sw, nameof(errorCaption), this.errorCaption);
            XmlHelper.WriteAttribute(sw, nameof(showError), this.showError);
            XmlHelper.WriteAttribute(sw, nameof(missingCaption), this.missingCaption);
            XmlHelper.WriteAttribute(sw, nameof(showMissing), this.showMissing);
            XmlHelper.WriteAttribute(sw, nameof(pageStyle), this.pageStyle);
            XmlHelper.WriteAttribute(sw, nameof(pivotTableStyle), this.pivotTableStyle);
            XmlHelper.WriteAttribute(sw, nameof(vacatedStyle), this.vacatedStyle);
            XmlHelper.WriteAttribute(sw, nameof(tag), this.tag);
            XmlHelper.WriteAttribute(sw, nameof(updatedVersion), this.updatedVersion);
            XmlHelper.WriteAttribute(sw, nameof(minRefreshableVersion), this.minRefreshableVersion);
            XmlHelper.WriteAttribute(sw, nameof(asteriskTotals), this.asteriskTotals);
            XmlHelper.WriteAttribute(sw, nameof(showItems), this.showItems);
            XmlHelper.WriteAttribute(sw, nameof(editData), this.editData);
            XmlHelper.WriteAttribute(sw, nameof(disableFieldList), this.disableFieldList);
            XmlHelper.WriteAttribute(sw, nameof(showCalcMbrs), this.showCalcMbrs);
            XmlHelper.WriteAttribute(sw, nameof(visualTotals), this.visualTotals);
            XmlHelper.WriteAttribute(sw, nameof(showMultipleLabel), this.showMultipleLabel);
            XmlHelper.WriteAttribute(sw, nameof(showDataDropDown), this.showDataDropDown);
            XmlHelper.WriteAttribute(sw, nameof(showDrill), this.showDrill);
            XmlHelper.WriteAttribute(sw, nameof(printDrill), this.printDrill);
            XmlHelper.WriteAttribute(sw, nameof(showMemberPropertyTips), this.showMemberPropertyTips);
            XmlHelper.WriteAttribute(sw, nameof(showDataTips), this.showDataTips);
            XmlHelper.WriteAttribute(sw, nameof(enableWizard), this.enableWizard);
            XmlHelper.WriteAttribute(sw, nameof(enableDrill), this.enableDrill);
            XmlHelper.WriteAttribute(sw, nameof(enableFieldProperties), this.enableFieldProperties);
            XmlHelper.WriteAttribute(sw, nameof(preserveFormatting), this.preserveFormatting);
            XmlHelper.WriteAttribute(sw, nameof(useAutoFormatting), this.useAutoFormatting);
            XmlHelper.WriteAttribute(sw, nameof(pageWrap), this.pageWrap);
            XmlHelper.WriteAttribute(sw, nameof(pageOverThenDown), this.pageOverThenDown);
            XmlHelper.WriteAttribute(sw, nameof(subtotalHiddenItems), this.subtotalHiddenItems);
            XmlHelper.WriteAttribute(sw, nameof(rowGrandTotals), this.rowGrandTotals);
            XmlHelper.WriteAttribute(sw, nameof(colGrandTotals), this.colGrandTotals);
            XmlHelper.WriteAttribute(sw, nameof(fieldPrintTitles), this.fieldPrintTitles);
            XmlHelper.WriteAttribute(sw, nameof(itemPrintTitles), this.itemPrintTitles);
            XmlHelper.WriteAttribute(sw, nameof(mergeItem), this.mergeItem);
            XmlHelper.WriteAttribute(sw, nameof(showDropZones), this.showDropZones);
            XmlHelper.WriteAttribute(sw, nameof(createdVersion), this.createdVersion);
            XmlHelper.WriteAttribute(sw, nameof(indent), this.indent);
            XmlHelper.WriteAttribute(sw, nameof(showEmptyRow), this.showEmptyRow);
            XmlHelper.WriteAttribute(sw, nameof(showEmptyCol), this.showEmptyCol);
            XmlHelper.WriteAttribute(sw, nameof(showHeaders), this.showHeaders);
            XmlHelper.WriteAttribute(sw, nameof(compact), this.compact);
            XmlHelper.WriteAttribute(sw, nameof(outline), this.outline);
            XmlHelper.WriteAttribute(sw, nameof(outlineData), this.outlineData);
            XmlHelper.WriteAttribute(sw, nameof(compactData), this.compactData);
            XmlHelper.WriteAttribute(sw, nameof(published), this.published);
            XmlHelper.WriteAttribute(sw, nameof(gridDropZones), this.gridDropZones);
            XmlHelper.WriteAttribute(sw, nameof(immersive), this.immersive);
            XmlHelper.WriteAttribute(sw, nameof(multipleFieldFilters), this.multipleFieldFilters);
            XmlHelper.WriteAttribute(sw, nameof(chartFormat), this.chartFormat);
            XmlHelper.WriteAttribute(sw, nameof(rowHeaderCaption), this.rowHeaderCaption);
            XmlHelper.WriteAttribute(sw, nameof(colHeaderCaption), this.colHeaderCaption);
            XmlHelper.WriteAttribute(sw, nameof(fieldListSortAscending), this.fieldListSortAscending);
            XmlHelper.WriteAttribute(sw, nameof(mdxSubqueries), this.mdxSubqueries);
            XmlHelper.WriteAttribute(sw, nameof(customListSort), this.customListSort);
            sw.Write(">");
            if (this.location != null)
                this.location.Write(sw, nameof(location));
            if (this.pivotFields != null)
                this.pivotFields.Write(sw, nameof(pivotFields));
            if (this.rowFields != null)
                this.rowFields.Write(sw, nameof(rowFields));
            if (this.rowItems != null)
                this.rowItems.Write(sw, nameof(rowItems));
            if (this.colFields != null)
                this.colFields.Write(sw, nameof(colFields));
            if (this.colItems != null)
                this.colItems.Write(sw, nameof(colItems));
            if (this.pageFields != null)
                this.pageFields.Write(sw, nameof(pageFields));
            if (this.dataFields != null)
                this.dataFields.Write(sw, nameof(dataFields));
            if (this.formats != null)
                this.formats.Write(sw, nameof(formats));
            if (this.conditionalFormats != null)
                this.conditionalFormats.Write(sw, nameof(conditionalFormats));
            if (this.chartFormats != null)
                this.chartFormats.Write(sw, nameof(chartFormats));
            if (this.pivotHierarchies != null)
                this.pivotHierarchies.Write(sw, nameof(pivotHierarchies));
            if (this.pivotTableStyleInfo != null)
                this.pivotTableStyleInfo.Write(sw, nameof(pivotTableStyleInfo));
            if (this.filters != null)
                this.filters.Write(sw, nameof(filters));
            if (this.rowHierarchiesUsage != null)
                this.rowHierarchiesUsage.Write(sw, nameof(rowHierarchiesUsage));
            if (this.colHierarchiesUsage != null)
                this.colHierarchiesUsage.Write(sw, nameof(colHierarchiesUsage));
            if (this.extLst != null)
                this.extLst.Write(sw, nameof(extLst));
            sw.Write(string.Format("</pivotTableDefinition>"));
        }
        public void Save(Stream stream)
        {
            using (StreamWriter sw = new StreamWriter(stream))
            {
                this.Write(sw);
            }
        }

        public CT_PivotTableDefinition()
        {
            //this.extLstField = new CT_ExtensionList();
            //this.colHierarchiesUsageField = new CT_ColHierarchiesUsage();
            //this.rowHierarchiesUsageField = new CT_RowHierarchiesUsage();
            //this.filtersField = new CT_PivotFilters();
            //this.pivotTableStyleInfoField = new CT_PivotTableStyle();
            //this.pivotHierarchiesField = new CT_PivotHierarchies();
            //this.chartFormatsField = new CT_ChartFormats();
            //this.conditionalFormatsField = new CT_ConditionalFormats();
            //this.formatsField = new CT_Formats();
            //this.dataFieldsField = new CT_DataFields();
            //this.pageFieldsField = new CT_PageFields();
            //this.colItemsField = new CT_colItems();
            //this.colFieldsField = new CT_ColFields();
            //this.rowItemsField = new CT_rowItems();
            //this.rowFieldsField = new CT_RowFields();
            //this.pivotFieldsField = new CT_PivotFields();
            //this.locationField = new CT_Location();
            this.dataOnRows = false;
            this.showError = false;
            this.showMissing = true;
            this.updatedVersion = ((byte)(0));
            this.minRefreshableVersion = ((byte)(0));
            this.asteriskTotals = false;
            this.showItems = true;
            this.editData = false;
            this.disableFieldList = false;
            this.showCalcMbrs = true;
            this.visualTotals = true;
            this.showMultipleLabel = true;
            this.showDataDropDown = true;
            this.showDrill = true;
            this.printDrill = false;
            this.showMemberPropertyTips = true;
            this.showDataTips = true;
            this.enableWizard = true;
            this.enableDrill = true;
            this.enableFieldProperties = true;
            this.preserveFormatting = true;
            this.useAutoFormatting = false;
            this.pageWrap = ((uint)(0));
            this.pageOverThenDown = false;
            this.subtotalHiddenItems = false;
            this.rowGrandTotals = true;
            this.colGrandTotals = true;
            this.fieldPrintTitles = false;
            this.itemPrintTitles = false;
            this.mergeItem = false;
            this.showDropZones = true;
            this.createdVersion = ((byte)(0));
            this.indent = ((uint)(1));
            this.showEmptyRow = false;
            this.showEmptyCol = false;
            this.showHeaders = true;
            this.compact = true;
            this.outline = false;
            this.outlineData = false;
            this.compactData = true;
            this.published = false;
            this.gridDropZones = false;
            this.immersive = true;
            this.multipleFieldFilters = true;
            this.chartFormat = ((uint)(0));
            this.fieldListSortAscending = false;
            this.mdxSubqueries = false;
            this.customListSort = true;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_Location location { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_PivotFields pivotFields { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 2)]
        public CT_RowFields rowFields { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 3)]
        public CT_rowItems rowItems { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 4)]
        public CT_ColFields colFields { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 5)]
        public CT_colItems colItems { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 6)]
        public CT_PageFields pageFields { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 7)]
        public CT_DataFields dataFields { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 8)]
        public CT_Formats formats { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 9)]
        public CT_ConditionalFormats conditionalFormats { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 10)]
        public CT_ChartFormats chartFormats { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 11)]
        public CT_PivotHierarchies pivotHierarchies { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 12)]
        public CT_PivotTableStyle pivotTableStyleInfo { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 13)]
        public CT_PivotFilters filters { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 14)]
        public CT_RowHierarchiesUsage rowHierarchiesUsage { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 15)]
        public CT_ColHierarchiesUsage colHierarchiesUsage { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 16)]
        public CT_ExtensionList extLst { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint cacheId { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool dataOnRows { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint dataPosition { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool dataPositionSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint autoFormatId { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool autoFormatIdSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyNumberFormats { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyNumberFormatsSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyBorderFormats { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyBorderFormatsSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyFontFormats { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyFontFormatsSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyPatternFormats { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyPatternFormatsSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyAlignmentFormats { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyAlignmentFormatsSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyWidthHeightFormats { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyWidthHeightFormatsSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string dataCaption { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string grandTotalCaption { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string errorCaption { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showError { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string missingCaption { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showMissing { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string pageStyle { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string pivotTableStyle { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string vacatedStyle { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string tag { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(byte), "0")]
        public byte updatedVersion { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(byte), "0")]
        public byte minRefreshableVersion { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool asteriskTotals { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showItems { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool editData { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool disableFieldList { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showCalcMbrs { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool visualTotals { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showMultipleLabel { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showDataDropDown { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showDrill { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool printDrill { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showMemberPropertyTips { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showDataTips { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool enableWizard { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool enableDrill { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool enableFieldProperties { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool preserveFormatting { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool useAutoFormatting { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint pageWrap { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool pageOverThenDown { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool subtotalHiddenItems { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool rowGrandTotals { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool colGrandTotals { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool fieldPrintTitles { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool itemPrintTitles { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool mergeItem { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showDropZones { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(byte), "0")]
        public byte createdVersion { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "1")]
        public uint indent { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showEmptyRow { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showEmptyCol { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showHeaders { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool compact { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool outline { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool outlineData { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool compactData { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool published { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool gridDropZones { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool immersive { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool multipleFieldFilters { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint chartFormat { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string rowHeaderCaption { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string colHeaderCaption { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool fieldListSortAscending { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool mdxSubqueries { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool customListSort { get; set; }


        public CT_PivotTableStyle AddNewPivotTableStyleInfo()
        {
            this.pivotTableStyleInfo = new CT_PivotTableStyle();
            return this.pivotTableStyleInfo;
        }

        public CT_RowFields AddNewRowFields()
        {
            this.rowFields = new CT_RowFields();
            return this.rowFields;
        }

        public CT_ColFields AddNewColFields()
        {
            this.colFields = new CT_ColFields();
            return this.colFields;
        }

        public CT_DataFields AddNewDataFields()
        {
            this.dataFields = new CT_DataFields();
            return this.dataFields;
        }

        public CT_PageFields AddNewPageFields()
        {
            this.pageFields = new CT_PageFields();
            return this.pageFields;
        }

        public CT_PivotFields AddNewPivotFields()
        {
            this.pivotFields = new CT_PivotFields();
            return this.pivotFields;
        }

        public CT_Location AddNewLocation()
        {
            this.location = new CT_Location();
            return this.location;
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Location
    {
        public static CT_Location Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Location ctObj = new CT_Location();
            ctObj.@ref = XmlHelper.ReadString(node.Attributes[nameof(@ref)]);
            if (node.Attributes[nameof(firstHeaderRow)] != null)
                ctObj.firstHeaderRow = XmlHelper.ReadUInt(node.Attributes[nameof(firstHeaderRow)]);
            if (node.Attributes[nameof(firstDataRow)] != null)
                ctObj.firstDataRow = XmlHelper.ReadUInt(node.Attributes[nameof(firstDataRow)]);
            if (node.Attributes[nameof(firstDataCol)] != null)
                ctObj.firstDataCol = XmlHelper.ReadUInt(node.Attributes[nameof(firstDataCol)]);
            if (node.Attributes[nameof(rowPageCount)] != null)
                ctObj.rowPageCount = XmlHelper.ReadUInt(node.Attributes[nameof(rowPageCount)]);
            if (node.Attributes[nameof(colPageCount)] != null)
                ctObj.colPageCount = XmlHelper.ReadUInt(node.Attributes[nameof(colPageCount)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(@ref), this.@ref);
            XmlHelper.WriteAttribute(sw, nameof(firstHeaderRow), this.firstHeaderRow);
            XmlHelper.WriteAttribute(sw, nameof(firstDataRow), this.firstDataRow);
            XmlHelper.WriteAttribute(sw, nameof(firstDataCol), this.firstDataCol);
            XmlHelper.WriteAttribute(sw, nameof(rowPageCount), this.rowPageCount);
            XmlHelper.WriteAttribute(sw, nameof(colPageCount), this.colPageCount);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Location()
        {
            this.rowPageCount = ((uint)(0));
            this.colPageCount = ((uint)(0));
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string @ref { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint firstHeaderRow { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint firstDataRow { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint firstDataCol { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint rowPageCount { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint colPageCount { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotFields
    {
        public static CT_PivotFields Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PivotFields ctObj = new CT_PivotFields();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.pivotField = new List<CT_PivotField>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(pivotField))
                    ctObj.pivotField.Add(CT_PivotField.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.pivotField != null)
            {
                foreach (CT_PivotField x in this.pivotField)
                {
                    x.Write(sw, nameof(pivotField));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_PivotFields()
        {
            this.pivotField = new List<CT_PivotField>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(pivotField), Order = 0)]
        public List<CT_PivotField> pivotField { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }

        public void SetPivotFieldArray(int columnIndex, CT_PivotField pivotField)
        {
            this.pivotField[columnIndex] = pivotField;
        }

        public CT_PivotField AddNewPivotField()
        {
            if (this.pivotField == null)
                this.pivotField = new List<CT_PivotField>();
            CT_PivotField f = new CT_PivotField();
            this.pivotField.Add(f);
            return f;
        }

        public uint SizeOfPivotFieldArray()
        {
            if (this.pivotField == null)
                this.pivotField = new List<CT_PivotField>();
            return (uint)this.pivotField.Count;
        }

        public CT_PivotField GetPivotFieldArray(int columnIndex)
        {
            if (this.pivotField == null)
                this.pivotField = new List<CT_PivotField>();
            return this.pivotField[columnIndex];
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotField
    {
        public static CT_PivotField Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PivotField ctObj = new CT_PivotField();
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            if (node.Attributes[nameof(axis)] != null)
                ctObj.axis = (ST_Axis)Enum.Parse(typeof(ST_Axis), node.Attributes[nameof(axis)].Value);
            if (node.Attributes[nameof(dataField)] != null)
                ctObj.dataField = XmlHelper.ReadBool(node.Attributes[nameof(dataField)]);
            ctObj.subtotalCaption = XmlHelper.ReadString(node.Attributes[nameof(subtotalCaption)]);
            if (node.Attributes[nameof(showDropDowns)] != null)
                ctObj.showDropDowns = XmlHelper.ReadBool(node.Attributes[nameof(showDropDowns)]);
            if (node.Attributes[nameof(hiddenLevel)] != null)
                ctObj.hiddenLevel = XmlHelper.ReadBool(node.Attributes[nameof(hiddenLevel)]);
            ctObj.uniqueMemberProperty = XmlHelper.ReadString(node.Attributes[nameof(uniqueMemberProperty)]);
            if (node.Attributes[nameof(compact)] != null)
                ctObj.compact = XmlHelper.ReadBool(node.Attributes[nameof(compact)]);
            if (node.Attributes[nameof(allDrilled)] != null)
                ctObj.allDrilled = XmlHelper.ReadBool(node.Attributes[nameof(allDrilled)]);
            if (node.Attributes[nameof(numFmtId)] != null)
                ctObj.numFmtId = XmlHelper.ReadUInt(node.Attributes[nameof(numFmtId)]);
            if (node.Attributes[nameof(outline)] != null)
                ctObj.outline = XmlHelper.ReadBool(node.Attributes[nameof(outline)]);
            if (node.Attributes[nameof(subtotalTop)] != null)
                ctObj.subtotalTop = XmlHelper.ReadBool(node.Attributes[nameof(subtotalTop)]);
            if (node.Attributes[nameof(dragToRow)] != null)
                ctObj.dragToRow = XmlHelper.ReadBool(node.Attributes[nameof(dragToRow)]);
            if (node.Attributes[nameof(dragToCol)] != null)
                ctObj.dragToCol = XmlHelper.ReadBool(node.Attributes[nameof(dragToCol)]);
            if (node.Attributes[nameof(multipleItemSelectionAllowed)] != null)
                ctObj.multipleItemSelectionAllowed = XmlHelper.ReadBool(node.Attributes[nameof(multipleItemSelectionAllowed)]);
            if (node.Attributes[nameof(dragToPage)] != null)
                ctObj.dragToPage = XmlHelper.ReadBool(node.Attributes[nameof(dragToPage)]);
            if (node.Attributes[nameof(dragToData)] != null)
                ctObj.dragToData = XmlHelper.ReadBool(node.Attributes[nameof(dragToData)]);
            if (node.Attributes[nameof(dragOff)] != null)
                ctObj.dragOff = XmlHelper.ReadBool(node.Attributes[nameof(dragOff)]);
            if (node.Attributes[nameof(showAll)] != null)
                ctObj.showAll = XmlHelper.ReadBool(node.Attributes[nameof(showAll)]);
            if (node.Attributes[nameof(insertBlankRow)] != null)
                ctObj.insertBlankRow = XmlHelper.ReadBool(node.Attributes[nameof(insertBlankRow)]);
            if (node.Attributes[nameof(serverField)] != null)
                ctObj.serverField = XmlHelper.ReadBool(node.Attributes[nameof(serverField)]);
            if (node.Attributes[nameof(insertPageBreak)] != null)
                ctObj.insertPageBreak = XmlHelper.ReadBool(node.Attributes[nameof(insertPageBreak)]);
            if (node.Attributes[nameof(autoShow)] != null)
                ctObj.autoShow = XmlHelper.ReadBool(node.Attributes[nameof(autoShow)]);
            if (node.Attributes[nameof(topAutoShow)] != null)
                ctObj.topAutoShow = XmlHelper.ReadBool(node.Attributes[nameof(topAutoShow)]);
            if (node.Attributes[nameof(hideNewItems)] != null)
                ctObj.hideNewItems = XmlHelper.ReadBool(node.Attributes[nameof(hideNewItems)]);
            if (node.Attributes[nameof(measureFilter)] != null)
                ctObj.measureFilter = XmlHelper.ReadBool(node.Attributes[nameof(measureFilter)]);
            if (node.Attributes[nameof(includeNewItemsInFilter)] != null)
                ctObj.includeNewItemsInFilter = XmlHelper.ReadBool(node.Attributes[nameof(includeNewItemsInFilter)]);
            if (node.Attributes[nameof(itemPageCount)] != null)
                ctObj.itemPageCount = XmlHelper.ReadUInt(node.Attributes[nameof(itemPageCount)]);
            if (node.Attributes[nameof(sortType)] != null)
                ctObj.sortType = (ST_FieldSortType)Enum.Parse(typeof(ST_FieldSortType), node.Attributes[nameof(sortType)].Value);
            if (node.Attributes[nameof(dataSourceSort)] != null)
                ctObj.dataSourceSort = XmlHelper.ReadBool(node.Attributes[nameof(dataSourceSort)]);
            if (node.Attributes[nameof(nonAutoSortDefault)] != null)
                ctObj.nonAutoSortDefault = XmlHelper.ReadBool(node.Attributes[nameof(nonAutoSortDefault)]);
            if (node.Attributes[nameof(rankBy)] != null)
                ctObj.rankBy = XmlHelper.ReadUInt(node.Attributes[nameof(rankBy)]);
            if (node.Attributes[nameof(defaultSubtotal)] != null)
                ctObj.defaultSubtotal = XmlHelper.ReadBool(node.Attributes[nameof(defaultSubtotal)]);
            if (node.Attributes[nameof(sumSubtotal)] != null)
                ctObj.sumSubtotal = XmlHelper.ReadBool(node.Attributes[nameof(sumSubtotal)]);
            if (node.Attributes[nameof(countASubtotal)] != null)
                ctObj.countASubtotal = XmlHelper.ReadBool(node.Attributes[nameof(countASubtotal)]);
            if (node.Attributes[nameof(avgSubtotal)] != null)
                ctObj.avgSubtotal = XmlHelper.ReadBool(node.Attributes[nameof(avgSubtotal)]);
            if (node.Attributes[nameof(maxSubtotal)] != null)
                ctObj.maxSubtotal = XmlHelper.ReadBool(node.Attributes[nameof(maxSubtotal)]);
            if (node.Attributes[nameof(minSubtotal)] != null)
                ctObj.minSubtotal = XmlHelper.ReadBool(node.Attributes[nameof(minSubtotal)]);
            if (node.Attributes[nameof(productSubtotal)] != null)
                ctObj.productSubtotal = XmlHelper.ReadBool(node.Attributes[nameof(productSubtotal)]);
            if (node.Attributes[nameof(countSubtotal)] != null)
                ctObj.countSubtotal = XmlHelper.ReadBool(node.Attributes[nameof(countSubtotal)]);
            if (node.Attributes[nameof(stdDevSubtotal)] != null)
                ctObj.stdDevSubtotal = XmlHelper.ReadBool(node.Attributes[nameof(stdDevSubtotal)]);
            if (node.Attributes[nameof(stdDevPSubtotal)] != null)
                ctObj.stdDevPSubtotal = XmlHelper.ReadBool(node.Attributes[nameof(stdDevPSubtotal)]);
            if (node.Attributes[nameof(varSubtotal)] != null)
                ctObj.varSubtotal = XmlHelper.ReadBool(node.Attributes[nameof(varSubtotal)]);
            if (node.Attributes[nameof(varPSubtotal)] != null)
                ctObj.varPSubtotal = XmlHelper.ReadBool(node.Attributes[nameof(varPSubtotal)]);
            if (node.Attributes[nameof(showPropCell)] != null)
                ctObj.showPropCell = XmlHelper.ReadBool(node.Attributes[nameof(showPropCell)]);
            if (node.Attributes[nameof(showPropTip)] != null)
                ctObj.showPropTip = XmlHelper.ReadBool(node.Attributes[nameof(showPropTip)]);
            if (node.Attributes[nameof(showPropAsCaption)] != null)
                ctObj.showPropAsCaption = XmlHelper.ReadBool(node.Attributes[nameof(showPropAsCaption)]);
            if (node.Attributes[nameof(defaultAttributeDrillState)] != null)
                ctObj.defaultAttributeDrillState = XmlHelper.ReadBool(node.Attributes[nameof(defaultAttributeDrillState)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(items))
                    ctObj.items = CT_Items.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(autoSortScope))
                    ctObj.autoSortScope = CT_AutoSortScope.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(axis), this.axis.ToString());
            XmlHelper.WriteAttribute(sw, nameof(dataField), this.dataField);
            XmlHelper.WriteAttribute(sw, nameof(subtotalCaption), this.subtotalCaption);
            XmlHelper.WriteAttribute(sw, nameof(showDropDowns), this.showDropDowns);
            XmlHelper.WriteAttribute(sw, nameof(hiddenLevel), this.hiddenLevel);
            XmlHelper.WriteAttribute(sw, nameof(uniqueMemberProperty), this.uniqueMemberProperty);
            XmlHelper.WriteAttribute(sw, nameof(compact), this.compact);
            XmlHelper.WriteAttribute(sw, nameof(allDrilled), this.allDrilled);
            XmlHelper.WriteAttribute(sw, nameof(numFmtId), this.numFmtId);
            XmlHelper.WriteAttribute(sw, nameof(outline), this.outline);
            XmlHelper.WriteAttribute(sw, nameof(subtotalTop), this.subtotalTop);
            XmlHelper.WriteAttribute(sw, nameof(dragToRow), this.dragToRow);
            XmlHelper.WriteAttribute(sw, nameof(dragToCol), this.dragToCol);
            XmlHelper.WriteAttribute(sw, nameof(multipleItemSelectionAllowed), this.multipleItemSelectionAllowed);
            XmlHelper.WriteAttribute(sw, nameof(dragToPage), this.dragToPage);
            XmlHelper.WriteAttribute(sw, nameof(dragToData), this.dragToData);
            XmlHelper.WriteAttribute(sw, nameof(dragOff), this.dragOff);
            XmlHelper.WriteAttribute(sw, nameof(showAll), this.showAll);
            XmlHelper.WriteAttribute(sw, nameof(insertBlankRow), this.insertBlankRow);
            XmlHelper.WriteAttribute(sw, nameof(serverField), this.serverField);
            XmlHelper.WriteAttribute(sw, nameof(insertPageBreak), this.insertPageBreak);
            XmlHelper.WriteAttribute(sw, nameof(autoShow), this.autoShow);
            XmlHelper.WriteAttribute(sw, nameof(topAutoShow), this.topAutoShow);
            XmlHelper.WriteAttribute(sw, nameof(hideNewItems), this.hideNewItems);
            XmlHelper.WriteAttribute(sw, nameof(measureFilter), this.measureFilter);
            XmlHelper.WriteAttribute(sw, nameof(includeNewItemsInFilter), this.includeNewItemsInFilter);
            XmlHelper.WriteAttribute(sw, nameof(itemPageCount), this.itemPageCount);
            XmlHelper.WriteAttribute(sw, nameof(sortType), this.sortType.ToString());
            XmlHelper.WriteAttribute(sw, nameof(dataSourceSort), this.dataSourceSort);
            XmlHelper.WriteAttribute(sw, nameof(nonAutoSortDefault), this.nonAutoSortDefault);
            XmlHelper.WriteAttribute(sw, nameof(rankBy), this.rankBy);
            XmlHelper.WriteAttribute(sw, nameof(defaultSubtotal), this.defaultSubtotal);
            XmlHelper.WriteAttribute(sw, nameof(sumSubtotal), this.sumSubtotal);
            XmlHelper.WriteAttribute(sw, nameof(countASubtotal), this.countASubtotal);
            XmlHelper.WriteAttribute(sw, nameof(avgSubtotal), this.avgSubtotal);
            XmlHelper.WriteAttribute(sw, nameof(maxSubtotal), this.maxSubtotal);
            XmlHelper.WriteAttribute(sw, nameof(minSubtotal), this.minSubtotal);
            XmlHelper.WriteAttribute(sw, nameof(productSubtotal), this.productSubtotal);
            XmlHelper.WriteAttribute(sw, nameof(countSubtotal), this.countSubtotal);
            XmlHelper.WriteAttribute(sw, nameof(stdDevSubtotal), this.stdDevSubtotal);
            XmlHelper.WriteAttribute(sw, nameof(stdDevPSubtotal), this.stdDevPSubtotal);
            XmlHelper.WriteAttribute(sw, nameof(varSubtotal), this.varSubtotal);
            XmlHelper.WriteAttribute(sw, nameof(varPSubtotal), this.varPSubtotal);
            XmlHelper.WriteAttribute(sw, nameof(showPropCell), this.showPropCell);
            XmlHelper.WriteAttribute(sw, nameof(showPropTip), this.showPropTip);
            XmlHelper.WriteAttribute(sw, nameof(showPropAsCaption), this.showPropAsCaption);
            XmlHelper.WriteAttribute(sw, nameof(defaultAttributeDrillState), this.defaultAttributeDrillState);
            sw.Write(">");
            if (this.items != null)
                this.items.Write(sw, nameof(items));
            if (this.autoSortScope != null)
                this.autoSortScope.Write(sw, nameof(autoSortScope));
            if (this.extLst != null)
                this.extLst.Write(sw, nameof(extLst));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        private bool showPropAsCaptionField;

        public CT_PivotField()
        {
            this.extLst = new CT_ExtensionList();
            this.autoSortScope = new CT_AutoSortScope();
            this.items = new CT_Items();
            this.dataField = false;
            this.showDropDowns = true;
            this.hiddenLevel = false;
            this.compact = true;
            this.allDrilled = false;
            this.outline = true;
            this.subtotalTop = true;
            this.dragToRow = true;
            this.dragToCol = true;
            this.multipleItemSelectionAllowed = false;
            this.dragToPage = true;
            this.dragToData = true;
            this.dragOff = true;
            this.showAll = true;
            this.insertBlankRow = false;
            this.serverField = false;
            this.insertPageBreak = false;
            this.autoShow = false;
            this.topAutoShow = true;
            this.hideNewItems = false;
            this.measureFilter = false;
            this.includeNewItemsInFilter = false;
            this.itemPageCount = ((uint)(10));
            this.sortType = ST_FieldSortType.manual;
            this.nonAutoSortDefault = false;
            this.defaultSubtotal = true;
            this.sumSubtotal = false;
            this.countASubtotal = false;
            this.avgSubtotal = false;
            this.maxSubtotal = false;
            this.minSubtotal = false;
            this.productSubtotal = false;
            this.countSubtotal = false;
            this.stdDevSubtotal = false;
            this.stdDevPSubtotal = false;
            this.varSubtotal = false;
            this.varPSubtotal = false;
            this.showPropCell = false;
            this.showPropTip = false;
            this.showPropAsCaptionField = false;
            this.defaultAttributeDrillState = false;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_Items items { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_AutoSortScope autoSortScope { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 2)]
        public CT_ExtensionList extLst { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_Axis axis { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool axisSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool dataField { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string subtotalCaption { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showDropDowns { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool hiddenLevel { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string uniqueMemberProperty { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool compact { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool allDrilled { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint numFmtId { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool numFmtIdSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool outline { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool subtotalTop { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToRow { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToCol { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool multipleItemSelectionAllowed { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToPage { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToData { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragOff { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showAll { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool insertBlankRow { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool serverField { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool insertPageBreak { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool autoShow { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool topAutoShow { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool hideNewItems { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool measureFilter { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool includeNewItemsInFilter { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "10")]
        public uint itemPageCount { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_FieldSortType.manual)]
        public ST_FieldSortType sortType { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool dataSourceSort { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool dataSourceSortSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool nonAutoSortDefault { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint rankBy { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool rankBySpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool defaultSubtotal { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool sumSubtotal { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool countASubtotal { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool avgSubtotal { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool maxSubtotal { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool minSubtotal { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool productSubtotal { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool countSubtotal { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool stdDevSubtotal { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool stdDevPSubtotal { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool varSubtotal { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool varPSubtotal { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showPropCell { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showPropTip { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showPropAsCaption
        {
            get
            {
                return this.showPropAsCaptionField;
            }
            set
            {
                this.showPropAsCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool defaultAttributeDrillState { get; set; }

        public CT_Items AddNewItems()
        {
            this.items = new CT_Items();
            return this.items;
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Items
    {
        public static CT_Items Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Items ctObj = new CT_Items();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.item = new List<CT_Item>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(item))
                    ctObj.item.Add(CT_Item.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.item != null)
            {
                foreach (CT_Item x in this.item)
                {
                    x.Write(sw, nameof(item));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Items()
        {
            this.item = new List<CT_Item>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(item), Order = 0)]
        public List<CT_Item> item { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }

        public CT_Item AddNewItem()
        {
            if (this.item == null)
                this.item = new List<CT_Item>();
            CT_Item i = new CT_Item();
            this.item.Add(i);
            return i;
        }

        public uint SizeOfItemArray()
        {
            if (this.item == null)
                this.item = new List<CT_Item>();
            return (uint)this.item.Count;
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Item
    {
        public static CT_Item Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Item ctObj = new CT_Item();
            ctObj.n = XmlHelper.ReadString(node.Attributes[nameof(n)]);
            if (node.Attributes[nameof(t)] != null)
                ctObj.t = (ST_ItemType)Enum.Parse(typeof(ST_ItemType), node.Attributes[nameof(t)].Value);
            if (node.Attributes[nameof(h)] != null)
                ctObj.h = XmlHelper.ReadBool(node.Attributes[nameof(h)]);
            if (node.Attributes[nameof(s)] != null)
                ctObj.s = XmlHelper.ReadBool(node.Attributes[nameof(s)]);
            if (node.Attributes[nameof(sd)] != null)
                ctObj.sd = XmlHelper.ReadBool(node.Attributes[nameof(sd)]);
            if (node.Attributes[nameof(f)] != null)
                ctObj.f = XmlHelper.ReadBool(node.Attributes[nameof(f)]);
            if (node.Attributes[nameof(m)] != null)
                ctObj.m = XmlHelper.ReadBool(node.Attributes[nameof(m)]);
            if (node.Attributes[nameof(c)] != null)
                ctObj.c = XmlHelper.ReadBool(node.Attributes[nameof(c)]);
            if (node.Attributes[nameof(x)] != null)
                ctObj.x = XmlHelper.ReadUInt(node.Attributes[nameof(x)]);
            if (node.Attributes[nameof(d)] != null)
                ctObj.d = XmlHelper.ReadBool(node.Attributes[nameof(d)]);
            if (node.Attributes[nameof(e)] != null)
                ctObj.e = XmlHelper.ReadBool(node.Attributes[nameof(e)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(n), this.n);
            XmlHelper.WriteAttribute(sw, nameof(t), this.t.ToString());
            XmlHelper.WriteAttribute(sw, nameof(h), this.h);
            XmlHelper.WriteAttribute(sw, nameof(s), this.s);
            XmlHelper.WriteAttribute(sw, nameof(sd), this.sd);
            XmlHelper.WriteAttribute(sw, nameof(f), this.f);
            XmlHelper.WriteAttribute(sw, nameof(m), this.m);
            XmlHelper.WriteAttribute(sw, nameof(c), this.c);
            XmlHelper.WriteAttribute(sw, nameof(x), this.x);
            XmlHelper.WriteAttribute(sw, nameof(d), this.d);
            XmlHelper.WriteAttribute(sw, nameof(e), this.e);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Item()
        {
            this.t = ST_ItemType.data;
            this.h = false;
            this.s = false;
            this.sd = true;
            this.f = false;
            this.m = false;
            this.c = false;
            this.d = false;
            this.e = true;
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string n { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_ItemType.data)]
        public ST_ItemType t { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool h { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool s { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool sd { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool f { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool m { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool c { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint x { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool xSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool d { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool e { get; set; }
    }

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_ItemType
    {

        /// <remarks/>
        data,

        /// <remarks/>
        @default,

        /// <remarks/>
        sum,

        /// <remarks/>
        countA,

        /// <remarks/>
        avg,

        /// <remarks/>
        max,

        /// <remarks/>
        min,

        /// <remarks/>
        product,

        /// <remarks/>
        count,

        /// <remarks/>
        stdDev,

        /// <remarks/>
        stdDevP,

        /// <remarks/>
        var,

        /// <remarks/>
        varP,

        /// <remarks/>
        grand,

        /// <remarks/>
        blank,
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_AutoSortScope
    {
        public static CT_AutoSortScope Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_AutoSortScope ctObj = new CT_AutoSortScope();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(pivotArea))
                    ctObj.pivotArea = CT_PivotArea.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            sw.Write(">");
            if (this.pivotArea != null)
                this.pivotArea.Write(sw, nameof(pivotArea));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_AutoSortScope()
        {
            this.pivotArea = new CT_PivotArea();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_PivotArea pivotArea { get; set; }
    }

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_FieldSortType
    {

        /// <remarks/>
        manual,

        /// <remarks/>
        ascending,

        /// <remarks/>
        descending,
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_RowFields
    {
        public static CT_RowFields Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RowFields ctObj = new CT_RowFields();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.field = new List<CT_Field>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(field))
                    ctObj.field.Add(CT_Field.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.field != null)
            {
                foreach (CT_Field x in this.field)
                {
                    x.Write(sw, nameof(field));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_RowFields()
        {
            this.field = new List<CT_Field>();
            this.count = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(field), Order = 0)]
        public List<CT_Field> field { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint count { get; set; }

        public CT_Field AddNewField()
        {
            CT_Field f = new CT_Field();
            this.field.Add(f);
            return f;
        }

        public uint SizeOfFieldArray()
        {
            return (uint)this.field.Count;
        }

        public List<CT_Field> GetFieldArray()
        {
            return this.field;
        }

        public CT_Field GetFieldArray(int p)
        {
            return this.field[p];
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Field
    {
        public static CT_Field Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Field ctObj = new CT_Field();
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int x { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_rowItems
    {
        public static CT_rowItems Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_rowItems ctObj = new CT_rowItems();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.i = new List<CT_I>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(i))
                    ctObj.i.Add(CT_I.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.i != null)
            {
                foreach (CT_I x in this.i)
                {
                    x.Write(sw, nameof(i));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_rowItems()
        {
            this.i = new List<CT_I>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(i), Order = 0)]
        public List<CT_I> i { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_I
    {
        public static CT_I Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_I ctObj = new CT_I();
            if (node.Attributes[nameof(t)] != null)
                ctObj.t = (ST_ItemType)Enum.Parse(typeof(ST_ItemType), node.Attributes[nameof(t)].Value);
            if (node.Attributes[nameof(r)] != null)
                ctObj.r = XmlHelper.ReadUInt(node.Attributes[nameof(r)]);
            if (node.Attributes[nameof(i)] != null)
                ctObj.i = XmlHelper.ReadUInt(node.Attributes[nameof(i)]);
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
            XmlHelper.WriteAttribute(sw, nameof(t), this.t.ToString());
            XmlHelper.WriteAttribute(sw, nameof(r), this.r);
            XmlHelper.WriteAttribute(sw, nameof(i), this.i);
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

        public CT_I()
        {
            this.x = new List<CT_X>();
            this.t = ST_ItemType.data;
            this.r = ((uint)(0));
            this.i = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(x), Order = 0)]
        public List<CT_X> x { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_ItemType.data)]
        public ST_ItemType t { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint r { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint i { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_ColFields
    {
        public static CT_ColFields Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ColFields ctObj = new CT_ColFields();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.field = new List<CT_Field>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(field))
                    ctObj.field.Add(CT_Field.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.field != null)
            {
                foreach (CT_Field x in this.field)
                {
                    x.Write(sw, nameof(field));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_ColFields()
        {
            this.field = new List<CT_Field>();
            this.count = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(field), Order = 0)]
        public List<CT_Field> field { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint count { get; set; }

        public uint SizeOfFieldArray()
        {
            if (this.field == null)
                this.field = new List<CT_Field>();
            return (uint)this.field.Count;
        }

        public CT_Field AddNewField()
        {
            if (this.field == null)
                this.field = new List<CT_Field>();
            CT_Field f = new CT_Field();
            this.field.Add(f);
            return f;
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_colItems
    {
        public static CT_colItems Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_colItems ctObj = new CT_colItems();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.i = new List<CT_I>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(i))
                    ctObj.i.Add(CT_I.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.i != null)
            {
                foreach (CT_I x in this.i)
                {
                    x.Write(sw, nameof(i));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_colItems()
        {
            this.i = new List<CT_I>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(i), Order = 0)]
        public List<CT_I> i { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PageFields
    {
        public static CT_PageFields Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PageFields ctObj = new CT_PageFields();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.pageField = new List<CT_PageField>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(pageField))
                    ctObj.pageField.Add(CT_PageField.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.pageField != null)
            {
                foreach (CT_PageField x in this.pageField)
                {
                    x.Write(sw, nameof(pageField));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_PageFields()
        {
            this.pageField = new List<CT_PageField>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(pageField), Order = 0)]
        public List<CT_PageField> pageField { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }

        public CT_PageField AddNewPageField()
        {
            if (this.pageField == null)
                this.pageField = new List<CT_PageField>();
            CT_PageField f = new CT_PageField();
            this.pageField.Add(f);
            return f;
        }

        public uint SizeOfPageFieldArray()
        {
            if (this.pageField == null)
                this.pageField = new List<CT_PageField>();
            return (uint)this.pageField.Count;
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PageField
    {
        public static CT_PageField Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PageField ctObj = new CT_PageField();
            if (node.Attributes[nameof(fld)] != null)
                ctObj.fld = XmlHelper.ReadInt(node.Attributes[nameof(fld)]);
            if (node.Attributes[nameof(item)] != null)
                ctObj.item = XmlHelper.ReadUInt(node.Attributes[nameof(item)]);
            if (node.Attributes[nameof(hier)] != null)
                ctObj.hier = XmlHelper.ReadInt(node.Attributes[nameof(hier)]);
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            ctObj.cap = XmlHelper.ReadString(node.Attributes[nameof(cap)]);
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
            XmlHelper.WriteAttribute(sw, nameof(fld), this.fld);
            XmlHelper.WriteAttribute(sw, nameof(item), this.item);
            XmlHelper.WriteAttribute(sw, nameof(hier), this.hier);
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(cap), this.cap);
            sw.Write(">");
            if (this.extLst != null)
                this.extLst.Write(sw, nameof(extLst));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_PageField()
        {
            this.extLst = new CT_ExtensionList();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_ExtensionList extLst { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int fld { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint item { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool itemSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int hier { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool hierSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string cap { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_DataFields
    {
        public static CT_DataFields Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_DataFields ctObj = new CT_DataFields();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.dataField = new List<CT_DataField>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(dataField))
                    ctObj.dataField.Add(CT_DataField.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.dataField != null)
            {
                foreach (CT_DataField x in this.dataField)
                {
                    x.Write(sw, nameof(dataField));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_DataFields()
        {
            this.dataField = new List<CT_DataField>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(dataField), Order = 0)]
        public List<CT_DataField> dataField { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }

        public CT_DataField AddNewDataField()
        {
            if (this.dataField == null)
                this.dataField = new List<CT_DataField>();
            CT_DataField f = new CT_DataField();
            this.dataField.Add(f);
            return f;
        }

        public uint SizeOfDataFieldArray()
        {
            return (uint)this.dataField.Count;
        }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_DataField
    {
        public static CT_DataField Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_DataField ctObj = new CT_DataField();
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            if (node.Attributes[nameof(fld)] != null)
                ctObj.fld = XmlHelper.ReadUInt(node.Attributes[nameof(fld)]);
            if (node.Attributes[nameof(subtotal)] != null)
                ctObj.subtotal = (ST_DataConsolidateFunction)Enum.Parse(typeof(ST_DataConsolidateFunction), node.Attributes[nameof(subtotal)].Value);
            if (node.Attributes[nameof(showDataAs)] != null)
                ctObj.showDataAs = (ST_ShowDataAs)Enum.Parse(typeof(ST_ShowDataAs), node.Attributes[nameof(showDataAs)].Value);
            if (node.Attributes[nameof(baseField)] != null)
                ctObj.baseField = XmlHelper.ReadInt(node.Attributes[nameof(baseField)]);
            if (node.Attributes[nameof(baseItem)] != null)
                ctObj.baseItem = XmlHelper.ReadUInt(node.Attributes[nameof(baseItem)]);
            if (node.Attributes[nameof(numFmtId)] != null)
                ctObj.numFmtId = XmlHelper.ReadUInt(node.Attributes[nameof(numFmtId)]);
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
            XmlHelper.WriteAttribute(sw, nameof(fld), this.fld);
            XmlHelper.WriteAttribute(sw, nameof(subtotal), this.subtotal.ToString());
            XmlHelper.WriteAttribute(sw, nameof(showDataAs), this.showDataAs.ToString());
            XmlHelper.WriteAttribute(sw, nameof(baseField), this.baseField);
            XmlHelper.WriteAttribute(sw, nameof(baseItem), this.baseItem);
            XmlHelper.WriteAttribute(sw, nameof(numFmtId), this.numFmtId);
            sw.Write(">");
            if (this.extLst != null)
                this.extLst.Write(sw, nameof(extLst));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_DataField()
        {
            this.extLst = new CT_ExtensionList();
            this.subtotal = ST_DataConsolidateFunction.sum;
            this.showDataAs = ST_ShowDataAs.normal;
            this.baseField = -1;
            this.baseItem = ((uint)(1048832));
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_ExtensionList extLst { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint fld { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_DataConsolidateFunction.sum)]
        public ST_DataConsolidateFunction subtotal { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_ShowDataAs.normal)]
        public ST_ShowDataAs showDataAs { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(-1)]
        public int baseField { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "1048832")]
        public uint baseItem { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint numFmtId { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool numFmtIdSpecified { get; set; }
    }

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_ShowDataAs
    {

        /// <remarks/>
        normal,

        /// <remarks/>
        difference,

        /// <remarks/>
        percent,

        /// <remarks/>
        percentDiff,

        /// <remarks/>
        runTotal,

        /// <remarks/>
        percentOfRow,

        /// <remarks/>
        percentOfCol,

        /// <remarks/>
        percentOfTotal,

        /// <remarks/>
        index,
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Formats
    {
        public static CT_Formats Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Formats ctObj = new CT_Formats();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.format = new List<CT_Format>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(format))
                    ctObj.format.Add(CT_Format.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.format != null)
            {
                foreach (CT_Format x in this.format)
                {
                    x.Write(sw, nameof(format));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Formats()
        {
            this.format = new List<CT_Format>();
            this.count = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(format), Order = 0)]
        public List<CT_Format> format { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint count { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Format
    {
        public static CT_Format Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Format ctObj = new CT_Format();
            if (node.Attributes[nameof(action)] != null)
                ctObj.action = (ST_FormatAction)Enum.Parse(typeof(ST_FormatAction), node.Attributes[nameof(action)].Value);
            if (node.Attributes[nameof(dxfId)] != null)
                ctObj.dxfId = XmlHelper.ReadUInt(node.Attributes[nameof(dxfId)]);
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
            XmlHelper.WriteAttribute(sw, nameof(action), this.action.ToString());
            XmlHelper.WriteAttribute(sw, nameof(dxfId), this.dxfId);
            sw.Write(">");
            if (this.pivotArea != null)
                this.pivotArea.Write(sw, nameof(pivotArea));
            if (this.extLst != null)
                this.extLst.Write(sw, nameof(extLst));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Format()
        {
            this.extLst = new CT_ExtensionList();
            this.pivotArea = new CT_PivotArea();
            this.action = ST_FormatAction.formatting;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_PivotArea pivotArea { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_ExtensionList extLst { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_FormatAction.formatting)]
        public ST_FormatAction action { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint dxfId { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool dxfIdSpecified { get; set; }
    }

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_FormatAction
    {

        /// <remarks/>
        blank,

        /// <remarks/>
        formatting,

        /// <remarks/>
        drill,

        /// <remarks/>
        formula,
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_ConditionalFormats
    {
        public static CT_ConditionalFormats Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ConditionalFormats ctObj = new CT_ConditionalFormats();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.conditionalFormat = new List<CT_ConditionalFormat>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(conditionalFormat))
                    ctObj.conditionalFormat.Add(CT_ConditionalFormat.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.conditionalFormat != null)
            {
                foreach (CT_ConditionalFormat x in this.conditionalFormat)
                {
                    x.Write(sw, nameof(conditionalFormat));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_ConditionalFormats()
        {
            this.conditionalFormat = new List<CT_ConditionalFormat>();
            this.count = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(conditionalFormat), Order = 0)]
        public List<CT_ConditionalFormat> conditionalFormat { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint count { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_ConditionalFormat
    {
        public static CT_ConditionalFormat Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ConditionalFormat ctObj = new CT_ConditionalFormat();
            if (node.Attributes[nameof(scope)] != null)
                ctObj.scope = (ST_Scope)Enum.Parse(typeof(ST_Scope), node.Attributes[nameof(scope)].Value);
            if (node.Attributes[nameof(type)] != null)
                ctObj.type = (ST_Type)Enum.Parse(typeof(ST_Type), node.Attributes[nameof(type)].Value);
            if (node.Attributes[nameof(priority)] != null)
                ctObj.priority = XmlHelper.ReadUInt(node.Attributes[nameof(priority)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(pivotAreas))
                    ctObj.pivotAreas = CT_PivotAreas.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(scope), this.scope.ToString());
            XmlHelper.WriteAttribute(sw, nameof(type), this.type.ToString());
            XmlHelper.WriteAttribute(sw, nameof(priority), this.priority);
            sw.Write(">");
            if (this.pivotAreas != null)
                this.pivotAreas.Write(sw, nameof(pivotAreas));
            if (this.extLst != null)
                this.extLst.Write(sw, nameof(extLst));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_ConditionalFormat()
        {
            this.extLst = new CT_ExtensionList();
            this.pivotAreas = new CT_PivotAreas();
            this.scope = ST_Scope.selection;
            this.type = ST_Type.none;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_PivotAreas pivotAreas { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_ExtensionList extLst { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_Scope.selection)]
        public ST_Scope scope { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_Type.none)]
        public ST_Type type { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint priority { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotAreas
    {
        public static CT_PivotAreas Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PivotAreas ctObj = new CT_PivotAreas();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.pivotArea = new List<CT_PivotArea>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(pivotArea))
                    ctObj.pivotArea.Add(CT_PivotArea.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.pivotArea != null)
            {
                foreach (CT_PivotArea x in this.pivotArea)
                {
                    x.Write(sw, nameof(pivotArea));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_PivotAreas()
        {
            this.pivotArea = new List<CT_PivotArea>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(pivotArea), Order = 0)]
        public List<CT_PivotArea> pivotArea { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_Scope
    {

        /// <remarks/>
        selection,

        /// <remarks/>
        data,

        /// <remarks/>
        field,
    }

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_Type
    {

        /// <remarks/>
        none,

        /// <remarks/>
        all,

        /// <remarks/>
        row,

        /// <remarks/>
        column,
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_ChartFormats
    {
        public static CT_ChartFormats Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ChartFormats ctObj = new CT_ChartFormats();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.chartFormat = new List<CT_ChartFormat>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(chartFormat))
                    ctObj.chartFormat.Add(CT_ChartFormat.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.chartFormat != null)
            {
                foreach (CT_ChartFormat x in this.chartFormat)
                {
                    x.Write(sw, nameof(chartFormat));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_ChartFormats()
        {
            this.chartFormat = new List<CT_ChartFormat>();
            this.count = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(chartFormat), Order = 0)]
        public List<CT_ChartFormat> chartFormat { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint count { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_ChartFormat
    {
        public static CT_ChartFormat Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ChartFormat ctObj = new CT_ChartFormat();
            if (node.Attributes[nameof(chart)] != null)
                ctObj.chart = XmlHelper.ReadUInt(node.Attributes[nameof(chart)]);
            if (node.Attributes[nameof(format)] != null)
                ctObj.format = XmlHelper.ReadUInt(node.Attributes[nameof(format)]);
            if (node.Attributes[nameof(series)] != null)
                ctObj.series = XmlHelper.ReadBool(node.Attributes[nameof(series)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(pivotArea))
                    ctObj.pivotArea = CT_PivotArea.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(chart), this.chart);
            XmlHelper.WriteAttribute(sw, nameof(format), this.format);
            XmlHelper.WriteAttribute(sw, nameof(series), this.series);
            sw.Write(">");
            if (this.pivotArea != null)
                this.pivotArea.Write(sw, nameof(pivotArea));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_ChartFormat()
        {
            this.pivotArea = new CT_PivotArea();
            this.series = false;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_PivotArea pivotArea { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint chart { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint format { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool series { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotHierarchies
    {
        public static CT_PivotHierarchies Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PivotHierarchies ctObj = new CT_PivotHierarchies();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.pivotHierarchy = new List<CT_PivotHierarchy>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(pivotHierarchy))
                    ctObj.pivotHierarchy.Add(CT_PivotHierarchy.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.pivotHierarchy != null)
            {
                foreach (CT_PivotHierarchy x in this.pivotHierarchy)
                {
                    x.Write(sw, nameof(pivotHierarchy));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_PivotHierarchies()
        {
            this.pivotHierarchy = new List<CT_PivotHierarchy>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(pivotHierarchy), Order = 0)]
        public List<CT_PivotHierarchy> pivotHierarchy { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotHierarchy
    {
        public static CT_PivotHierarchy Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PivotHierarchy ctObj = new CT_PivotHierarchy();
            if (node.Attributes[nameof(outline)] != null)
                ctObj.outline = XmlHelper.ReadBool(node.Attributes[nameof(outline)]);
            if (node.Attributes[nameof(multipleItemSelectionAllowed)] != null)
                ctObj.multipleItemSelectionAllowed = XmlHelper.ReadBool(node.Attributes[nameof(multipleItemSelectionAllowed)]);
            if (node.Attributes[nameof(subtotalTop)] != null)
                ctObj.subtotalTop = XmlHelper.ReadBool(node.Attributes[nameof(subtotalTop)]);
            if (node.Attributes[nameof(showInFieldList)] != null)
                ctObj.showInFieldList = XmlHelper.ReadBool(node.Attributes[nameof(showInFieldList)]);
            if (node.Attributes[nameof(dragToRow)] != null)
                ctObj.dragToRow = XmlHelper.ReadBool(node.Attributes[nameof(dragToRow)]);
            if (node.Attributes[nameof(dragToCol)] != null)
                ctObj.dragToCol = XmlHelper.ReadBool(node.Attributes[nameof(dragToCol)]);
            if (node.Attributes[nameof(dragToPage)] != null)
                ctObj.dragToPage = XmlHelper.ReadBool(node.Attributes[nameof(dragToPage)]);
            if (node.Attributes[nameof(dragToData)] != null)
                ctObj.dragToData = XmlHelper.ReadBool(node.Attributes[nameof(dragToData)]);
            if (node.Attributes[nameof(dragOff)] != null)
                ctObj.dragOff = XmlHelper.ReadBool(node.Attributes[nameof(dragOff)]);
            if (node.Attributes[nameof(includeNewItemsInFilter)] != null)
                ctObj.includeNewItemsInFilter = XmlHelper.ReadBool(node.Attributes[nameof(includeNewItemsInFilter)]);
            ctObj.caption = XmlHelper.ReadString(node.Attributes[nameof(caption)]);
            ctObj.members = new List<CT_Members>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(mps))
                    ctObj.mps = CT_MemberProperties.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(members))
                    ctObj.members.Add(CT_Members.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(outline), this.outline);
            XmlHelper.WriteAttribute(sw, nameof(multipleItemSelectionAllowed), this.multipleItemSelectionAllowed);
            XmlHelper.WriteAttribute(sw, nameof(subtotalTop), this.subtotalTop);
            XmlHelper.WriteAttribute(sw, nameof(showInFieldList), this.showInFieldList);
            XmlHelper.WriteAttribute(sw, nameof(dragToRow), this.dragToRow);
            XmlHelper.WriteAttribute(sw, nameof(dragToCol), this.dragToCol);
            XmlHelper.WriteAttribute(sw, nameof(dragToPage), this.dragToPage);
            XmlHelper.WriteAttribute(sw, nameof(dragToData), this.dragToData);
            XmlHelper.WriteAttribute(sw, nameof(dragOff), this.dragOff);
            XmlHelper.WriteAttribute(sw, nameof(includeNewItemsInFilter), this.includeNewItemsInFilter);
            XmlHelper.WriteAttribute(sw, nameof(caption), this.caption);
            sw.Write(">");
            if (this.mps != null)
                this.mps.Write(sw, nameof(mps));
            if (this.extLst != null)
                this.extLst.Write(sw, nameof(extLst));
            if (this.members != null)
            {
                foreach (CT_Members x in this.members)
                {
                    x.Write(sw, nameof(members));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_PivotHierarchy()
        {
            this.extLst = new CT_ExtensionList();
            this.members = new List<CT_Members>();
            this.mps = new CT_MemberProperties();
            this.outline = false;
            this.multipleItemSelectionAllowed = false;
            this.subtotalTop = false;
            this.showInFieldList = true;
            this.dragToRow = true;
            this.dragToCol = true;
            this.dragToPage = true;
            this.dragToData = false;
            this.dragOff = true;
            this.includeNewItemsInFilter = false;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_MemberProperties mps { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(nameof(members), Order = 1)]
        public List<CT_Members> members { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 2)]
        public CT_ExtensionList extLst { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool outline { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool multipleItemSelectionAllowed { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool subtotalTop { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showInFieldList { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToRow { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToCol { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToPage { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool dragToData { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragOff { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool includeNewItemsInFilter { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string caption { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_MemberProperties
    {
        public static CT_MemberProperties Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MemberProperties ctObj = new CT_MemberProperties();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.mp = new List<CT_MemberProperty>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(mp))
                    ctObj.mp.Add(CT_MemberProperty.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.mp != null)
            {
                foreach (CT_MemberProperty x in this.mp)
                {
                    x.Write(sw, nameof(mp));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_MemberProperties()
        {
            this.mp = new List<CT_MemberProperty>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(mp), Order = 0)]
        public List<CT_MemberProperty> mp { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_MemberProperty
    {
        public static CT_MemberProperty Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_MemberProperty ctObj = new CT_MemberProperty();
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            if (node.Attributes[nameof(showCell)] != null)
                ctObj.showCell = XmlHelper.ReadBool(node.Attributes[nameof(showCell)]);
            if (node.Attributes[nameof(showTip)] != null)
                ctObj.showTip = XmlHelper.ReadBool(node.Attributes[nameof(showTip)]);
            if (node.Attributes[nameof(showAsCaption)] != null)
                ctObj.showAsCaption = XmlHelper.ReadBool(node.Attributes[nameof(showAsCaption)]);
            if (node.Attributes[nameof(nameLen)] != null)
                ctObj.nameLen = XmlHelper.ReadUInt(node.Attributes[nameof(nameLen)]);
            if (node.Attributes[nameof(pPos)] != null)
                ctObj.pPos = XmlHelper.ReadUInt(node.Attributes[nameof(pPos)]);
            if (node.Attributes[nameof(pLen)] != null)
                ctObj.pLen = XmlHelper.ReadUInt(node.Attributes[nameof(pLen)]);
            if (node.Attributes[nameof(level)] != null)
                ctObj.level = XmlHelper.ReadUInt(node.Attributes[nameof(level)]);
            if (node.Attributes[nameof(field)] != null)
                ctObj.field = XmlHelper.ReadUInt(node.Attributes[nameof(field)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(showCell), this.showCell);
            XmlHelper.WriteAttribute(sw, nameof(showTip), this.showTip);
            XmlHelper.WriteAttribute(sw, nameof(showAsCaption), this.showAsCaption);
            XmlHelper.WriteAttribute(sw, nameof(nameLen), this.nameLen);
            XmlHelper.WriteAttribute(sw, nameof(pPos), this.pPos);
            XmlHelper.WriteAttribute(sw, nameof(pLen), this.pLen);
            XmlHelper.WriteAttribute(sw, nameof(level), this.level);
            XmlHelper.WriteAttribute(sw, nameof(field), this.field);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_MemberProperty()
        {
            this.showCell = false;
            this.showTip = false;
            this.showAsCaption = false;
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showCell { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showTip { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showAsCaption { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint nameLen { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool nameLenSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint pPos { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool pPosSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint pLen { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool pLenSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint level { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool levelSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint field { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Members
    {
        public static CT_Members Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Members ctObj = new CT_Members();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            if (node.Attributes[nameof(level)] != null)
                ctObj.level = XmlHelper.ReadUInt(node.Attributes[nameof(level)]);
            ctObj.member = new List<CT_Member>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(member))
                    ctObj.member.Add(CT_Member.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            XmlHelper.WriteAttribute(sw, nameof(level), this.level);
            sw.Write(">");
            if (this.member != null)
            {
                foreach (CT_Member x in this.member)
                {
                    x.Write(sw, nameof(member));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_Members()
        {
            this.member = new List<CT_Member>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(member), Order = 0)]
        public List<CT_Member> member { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint level { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool levelSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_Member
    {
        public static CT_Member Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_Member ctObj = new CT_Member();
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

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotTableStyle
    {
        public static CT_PivotTableStyle Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PivotTableStyle ctObj = new CT_PivotTableStyle();
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            if (node.Attributes[nameof(showRowHeaders)] != null)
                ctObj.showRowHeaders = XmlHelper.ReadBool(node.Attributes[nameof(showRowHeaders)]);
            if (node.Attributes[nameof(showColHeaders)] != null)
                ctObj.showColHeaders = XmlHelper.ReadBool(node.Attributes[nameof(showColHeaders)]);
            if (node.Attributes[nameof(showRowStripes)] != null)
                ctObj.showRowStripes = XmlHelper.ReadBool(node.Attributes[nameof(showRowStripes)]);
            if (node.Attributes[nameof(showColStripes)] != null)
                ctObj.showColStripes = XmlHelper.ReadBool(node.Attributes[nameof(showColStripes)]);
            if (node.Attributes[nameof(showLastColumn)] != null)
                ctObj.showLastColumn = XmlHelper.ReadBool(node.Attributes[nameof(showLastColumn)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(showRowHeaders), this.showRowHeaders);
            XmlHelper.WriteAttribute(sw, nameof(showColHeaders), this.showColHeaders);
            XmlHelper.WriteAttribute(sw, nameof(showRowStripes), this.showRowStripes);
            XmlHelper.WriteAttribute(sw, nameof(showColStripes), this.showColStripes);
            XmlHelper.WriteAttribute(sw, nameof(showLastColumn), this.showLastColumn);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool showRowHeaders { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool showRowHeadersSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool showColHeaders { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool showColHeadersSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool showRowStripes { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool showRowStripesSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool showColStripes { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool showColStripesSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool showLastColumn { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool showLastColumnSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotFilters
    {
        public static CT_PivotFilters Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PivotFilters ctObj = new CT_PivotFilters();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.filter = new List<CT_PivotFilter>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(filter))
                    ctObj.filter.Add(CT_PivotFilter.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.filter != null)
            {
                foreach (CT_PivotFilter x in this.filter)
                {
                    x.Write(sw, nameof(filter));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_PivotFilters()
        {
            this.filter = new List<CT_PivotFilter>();
            this.count = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(filter), Order = 0)]
        public List<CT_PivotFilter> filter { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint count { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_PivotFilter
    {
        public static CT_PivotFilter Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_PivotFilter ctObj = new CT_PivotFilter();
            if (node.Attributes[nameof(fld)] != null)
                ctObj.fld = XmlHelper.ReadUInt(node.Attributes[nameof(fld)]);
            if (node.Attributes[nameof(mpFld)] != null)
                ctObj.mpFld = XmlHelper.ReadUInt(node.Attributes[nameof(mpFld)]);
            if (node.Attributes[nameof(type)] != null)
                ctObj.type = (ST_PivotFilterType)Enum.Parse(typeof(ST_PivotFilterType), node.Attributes[nameof(type)].Value);
            if (node.Attributes[nameof(evalOrder)] != null)
                ctObj.evalOrder = XmlHelper.ReadInt(node.Attributes[nameof(evalOrder)]);
            if (node.Attributes[nameof(id)] != null)
                ctObj.id = XmlHelper.ReadUInt(node.Attributes[nameof(id)]);
            if (node.Attributes[nameof(iMeasureHier)] != null)
                ctObj.iMeasureHier = XmlHelper.ReadUInt(node.Attributes[nameof(iMeasureHier)]);
            if (node.Attributes[nameof(iMeasureFld)] != null)
                ctObj.iMeasureFld = XmlHelper.ReadUInt(node.Attributes[nameof(iMeasureFld)]);
            ctObj.name = XmlHelper.ReadString(node.Attributes[nameof(name)]);
            ctObj.description = XmlHelper.ReadString(node.Attributes[nameof(description)]);
            ctObj.stringValue1 = XmlHelper.ReadString(node.Attributes[nameof(stringValue1)]);
            ctObj.stringValue2 = XmlHelper.ReadString(node.Attributes[nameof(stringValue2)]);
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(autoFilter))
                    ctObj.autoFilter = CT_AutoFilter.Parse(childNode, namespaceManager);
                else if (childNode.LocalName == nameof(extLst))
                    ctObj.extLst = CT_ExtensionList.Parse(childNode, namespaceManager);
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(fld), this.fld);
            XmlHelper.WriteAttribute(sw, nameof(mpFld), this.mpFld);
            XmlHelper.WriteAttribute(sw, nameof(type), this.type.ToString());
            XmlHelper.WriteAttribute(sw, nameof(evalOrder), this.evalOrder);
            XmlHelper.WriteAttribute(sw, nameof(id), this.id);
            XmlHelper.WriteAttribute(sw, nameof(iMeasureHier), this.iMeasureHier);
            XmlHelper.WriteAttribute(sw, nameof(iMeasureFld), this.iMeasureFld);
            XmlHelper.WriteAttribute(sw, nameof(name), this.name);
            XmlHelper.WriteAttribute(sw, nameof(description), this.description);
            XmlHelper.WriteAttribute(sw, nameof(stringValue1), this.stringValue1);
            XmlHelper.WriteAttribute(sw, nameof(stringValue2), this.stringValue2);
            sw.Write(">");
            if (this.autoFilter != null)
                this.autoFilter.Write(sw, nameof(autoFilter));
            if (this.extLst != null)
                this.extLst.Write(sw, nameof(extLst));
            sw.Write(string.Format("</{0}>", nodeName));
        }

        private string stringValue2Field;

        public CT_PivotFilter()
        {
            this.extLst = new CT_ExtensionList();
            this.autoFilter = new CT_AutoFilter();
            this.evalOrder = 0;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_AutoFilter autoFilter { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_ExtensionList extLst { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint fld { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint mpFld { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool mpFldSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_PivotFilterType type { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public int evalOrder { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint id { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint iMeasureHier { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool iMeasureHierSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint iMeasureFld { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool iMeasureFldSpecified { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string description { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string stringValue1 { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string stringValue2
        {
            get
            {
                return this.stringValue2Field;
            }
            set
            {
                this.stringValue2Field = value;
            }
        }
    }

    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public enum ST_PivotFilterType
    {

        /// <remarks/>
        unknown,

        /// <remarks/>
        count,

        /// <remarks/>
        percent,

        /// <remarks/>
        sum,

        /// <remarks/>
        captionEqual,

        /// <remarks/>
        captionNotEqual,

        /// <remarks/>
        captionBeginsWith,

        /// <remarks/>
        captionNotBeginsWith,

        /// <remarks/>
        captionEndsWith,

        /// <remarks/>
        captionNotEndsWith,

        /// <remarks/>
        captionContains,

        /// <remarks/>
        captionNotContains,

        /// <remarks/>
        captionGreaterThan,

        /// <remarks/>
        captionGreaterThanOrEqual,

        /// <remarks/>
        captionLessThan,

        /// <remarks/>
        captionLessThanOrEqual,

        /// <remarks/>
        captionBetween,

        /// <remarks/>
        captionNotBetween,

        /// <remarks/>
        valueEqual,

        /// <remarks/>
        valueNotEqual,

        /// <remarks/>
        valueGreaterThan,

        /// <remarks/>
        valueGreaterThanOrEqual,

        /// <remarks/>
        valueLessThan,

        /// <remarks/>
        valueLessThanOrEqual,

        /// <remarks/>
        valueBetween,

        /// <remarks/>
        valueNotBetween,

        /// <remarks/>
        dateEqual,

        /// <remarks/>
        dateNotEqual,

        /// <remarks/>
        dateOlderThan,

        /// <remarks/>
        dateOlderThanOrEqual,

        /// <remarks/>
        dateNewerThan,

        /// <remarks/>
        dateNewerThanOrEqual,

        /// <remarks/>
        dateBetween,

        /// <remarks/>
        dateNotBetween,

        /// <remarks/>
        tomorrow,

        /// <remarks/>
        today,

        /// <remarks/>
        yesterday,

        /// <remarks/>
        nextWeek,

        /// <remarks/>
        thisWeek,

        /// <remarks/>
        lastWeek,

        /// <remarks/>
        nextMonth,

        /// <remarks/>
        thisMonth,

        /// <remarks/>
        lastMonth,

        /// <remarks/>
        nextQuarter,

        /// <remarks/>
        thisQuarter,

        /// <remarks/>
        lastQuarter,

        /// <remarks/>
        nextYear,

        /// <remarks/>
        thisYear,

        /// <remarks/>
        lastYear,

        /// <remarks/>
        yearToDate,

        /// <remarks/>
        Q1,

        /// <remarks/>
        Q2,

        /// <remarks/>
        Q3,

        /// <remarks/>
        Q4,

        /// <remarks/>
        M1,

        /// <remarks/>
        M2,

        /// <remarks/>
        M3,

        /// <remarks/>
        M4,

        /// <remarks/>
        M5,

        /// <remarks/>
        M6,

        /// <remarks/>
        M7,

        /// <remarks/>
        M8,

        /// <remarks/>
        M9,

        /// <remarks/>
        M10,

        /// <remarks/>
        M11,

        /// <remarks/>
        M12,
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_RowHierarchiesUsage
    {
        public static CT_RowHierarchiesUsage Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_RowHierarchiesUsage ctObj = new CT_RowHierarchiesUsage();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.rowHierarchyUsage = new List<CT_HierarchyUsage>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(rowHierarchyUsage))
                    ctObj.rowHierarchyUsage.Add(CT_HierarchyUsage.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.rowHierarchyUsage != null)
            {
                foreach (CT_HierarchyUsage x in this.rowHierarchyUsage)
                {
                    x.Write(sw, nameof(rowHierarchyUsage));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_RowHierarchiesUsage()
        {
            this.rowHierarchyUsage = new List<CT_HierarchyUsage>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(rowHierarchyUsage), Order = 0)]
        public List<CT_HierarchyUsage> rowHierarchyUsage { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_HierarchyUsage
    {
        public static CT_HierarchyUsage Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_HierarchyUsage ctObj = new CT_HierarchyUsage();
            if (node.Attributes[nameof(hierarchyUsage)] != null)
                ctObj.hierarchyUsage = XmlHelper.ReadInt(node.Attributes[nameof(hierarchyUsage)]);
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(hierarchyUsage), this.hierarchyUsage);
            sw.Write(">");
            sw.Write(string.Format("</{0}>", nodeName));
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int hierarchyUsage { get; set; }
    }

    
    
    
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = true)]
    public partial class CT_ColHierarchiesUsage
    {
        public static CT_ColHierarchiesUsage Parse(XmlNode node, XmlNamespaceManager namespaceManager)
        {
            if (node == null)
                return null;
            CT_ColHierarchiesUsage ctObj = new CT_ColHierarchiesUsage();
            if (node.Attributes[nameof(count)] != null)
                ctObj.count = XmlHelper.ReadUInt(node.Attributes[nameof(count)]);
            ctObj.colHierarchyUsage = new List<CT_HierarchyUsage>();
            foreach (XmlNode childNode in node.ChildNodes)
            {
                if (childNode.LocalName == nameof(colHierarchyUsage))
                    ctObj.colHierarchyUsage.Add(CT_HierarchyUsage.Parse(childNode, namespaceManager));
            }
            return ctObj;
        }



        internal void Write(StreamWriter sw, string nodeName)
        {
            sw.Write(string.Format("<{0}", nodeName));
            XmlHelper.WriteAttribute(sw, nameof(count), this.count);
            sw.Write(">");
            if (this.colHierarchyUsage != null)
            {
                foreach (CT_HierarchyUsage x in this.colHierarchyUsage)
                {
                    x.Write(sw, nameof(colHierarchyUsage));
                }
            }
            sw.Write(string.Format("</{0}>", nodeName));
        }

        public CT_ColHierarchiesUsage()
        {
            this.colHierarchyUsage = new List<CT_HierarchyUsage>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(colHierarchyUsage), Order = 0)]
        public List<CT_HierarchyUsage> colHierarchyUsage { get; set; }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified { get; set; }
    }
    
}
