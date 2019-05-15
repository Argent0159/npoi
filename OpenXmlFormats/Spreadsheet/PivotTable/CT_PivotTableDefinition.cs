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
        private CT_Location locationField;

        private CT_PivotFields pivotFieldsField;

        private CT_RowFields rowFieldsField;

        private CT_rowItems rowItemsField;

        private CT_ColFields colFieldsField;

        private CT_colItems colItemsField;

        private CT_PageFields pageFieldsField;

        private CT_DataFields dataFieldsField;

        private CT_Formats formatsField;

        private CT_ConditionalFormats conditionalFormatsField;

        private CT_ChartFormats chartFormatsField;

        private CT_PivotHierarchies pivotHierarchiesField;

        private CT_PivotTableStyle pivotTableStyleInfoField;

        private CT_PivotFilters filtersField;

        private CT_RowHierarchiesUsage rowHierarchiesUsageField;

        private CT_ColHierarchiesUsage colHierarchiesUsageField;

        private CT_ExtensionList extLstField;

        private string nameField;

        private uint cacheIdField;

        private bool dataOnRowsField;

        private uint dataPositionField;

        private bool dataPositionFieldSpecified;

        private uint autoFormatIdField;

        private bool autoFormatIdFieldSpecified;

        private bool applyNumberFormatsField;

        private bool applyNumberFormatsFieldSpecified;

        private bool applyBorderFormatsField;

        private bool applyBorderFormatsFieldSpecified;

        private bool applyFontFormatsField;

        private bool applyFontFormatsFieldSpecified;

        private bool applyPatternFormatsField;

        private bool applyPatternFormatsFieldSpecified;

        private bool applyAlignmentFormatsField;

        private bool applyAlignmentFormatsFieldSpecified;

        private bool applyWidthHeightFormatsField;

        private bool applyWidthHeightFormatsFieldSpecified;

        private string dataCaptionField;

        private string grandTotalCaptionField;

        private string errorCaptionField;

        private bool showErrorField;

        private string missingCaptionField;

        private bool showMissingField;

        private string pageStyleField;

        private string pivotTableStyleField;

        private string vacatedStyleField;

        private string tagField;

        private byte updatedVersionField;

        private byte minRefreshableVersionField;

        private bool asteriskTotalsField;

        private bool showItemsField;

        private bool editDataField;

        private bool disableFieldListField;

        private bool showCalcMbrsField;

        private bool visualTotalsField;

        private bool showMultipleLabelField;

        private bool showDataDropDownField;

        private bool showDrillField;

        private bool printDrillField;

        private bool showMemberPropertyTipsField;

        private bool showDataTipsField;

        private bool enableWizardField;

        private bool enableDrillField;

        private bool enableFieldPropertiesField;

        private bool preserveFormattingField;

        private bool useAutoFormattingField;

        private uint pageWrapField;

        private bool pageOverThenDownField;

        private bool subtotalHiddenItemsField;

        private bool rowGrandTotalsField;

        private bool colGrandTotalsField;

        private bool fieldPrintTitlesField;

        private bool itemPrintTitlesField;

        private bool mergeItemField;

        private bool showDropZonesField;

        private byte createdVersionField;

        private uint indentField;

        private bool showEmptyRowField;

        private bool showEmptyColField;

        private bool showHeadersField;

        private bool compactField;

        private bool outlineField;

        private bool outlineDataField;

        private bool compactDataField;

        private bool publishedField;

        private bool gridDropZonesField;

        private bool immersiveField;

        private bool multipleFieldFiltersField;

        private uint chartFormatField;

        private string rowHeaderCaptionField;

        private string colHeaderCaptionField;

        private bool fieldListSortAscendingField;

        private bool mdxSubqueriesField;

        private bool customListSortField;

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
            this.dataOnRowsField = false;
            this.showErrorField = false;
            this.showMissingField = true;
            this.updatedVersionField = ((byte)(0));
            this.minRefreshableVersionField = ((byte)(0));
            this.asteriskTotalsField = false;
            this.showItemsField = true;
            this.editDataField = false;
            this.disableFieldListField = false;
            this.showCalcMbrsField = true;
            this.visualTotalsField = true;
            this.showMultipleLabelField = true;
            this.showDataDropDownField = true;
            this.showDrillField = true;
            this.printDrillField = false;
            this.showMemberPropertyTipsField = true;
            this.showDataTipsField = true;
            this.enableWizardField = true;
            this.enableDrillField = true;
            this.enableFieldPropertiesField = true;
            this.preserveFormattingField = true;
            this.useAutoFormattingField = false;
            this.pageWrapField = ((uint)(0));
            this.pageOverThenDownField = false;
            this.subtotalHiddenItemsField = false;
            this.rowGrandTotalsField = true;
            this.colGrandTotalsField = true;
            this.fieldPrintTitlesField = false;
            this.itemPrintTitlesField = false;
            this.mergeItemField = false;
            this.showDropZonesField = true;
            this.createdVersionField = ((byte)(0));
            this.indentField = ((uint)(1));
            this.showEmptyRowField = false;
            this.showEmptyColField = false;
            this.showHeadersField = true;
            this.compactField = true;
            this.outlineField = false;
            this.outlineDataField = false;
            this.compactDataField = true;
            this.publishedField = false;
            this.gridDropZonesField = false;
            this.immersiveField = true;
            this.multipleFieldFiltersField = true;
            this.chartFormatField = ((uint)(0));
            this.fieldListSortAscendingField = false;
            this.mdxSubqueriesField = false;
            this.customListSortField = true;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_Location location
        {
            get
            {
                return this.locationField;
            }
            set
            {
                this.locationField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_PivotFields pivotFields
        {
            get
            {
                return this.pivotFieldsField;
            }
            set
            {
                this.pivotFieldsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 2)]
        public CT_RowFields rowFields
        {
            get
            {
                return this.rowFieldsField;
            }
            set
            {
                this.rowFieldsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 3)]
        public CT_rowItems rowItems
        {
            get
            {
                return this.rowItemsField;
            }
            set
            {
                this.rowItemsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 4)]
        public CT_ColFields colFields
        {
            get
            {
                return this.colFieldsField;
            }
            set
            {
                this.colFieldsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 5)]
        public CT_colItems colItems
        {
            get
            {
                return this.colItemsField;
            }
            set
            {
                this.colItemsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 6)]
        public CT_PageFields pageFields
        {
            get
            {
                return this.pageFieldsField;
            }
            set
            {
                this.pageFieldsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 7)]
        public CT_DataFields dataFields
        {
            get
            {
                return this.dataFieldsField;
            }
            set
            {
                this.dataFieldsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 8)]
        public CT_Formats formats
        {
            get
            {
                return this.formatsField;
            }
            set
            {
                this.formatsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 9)]
        public CT_ConditionalFormats conditionalFormats
        {
            get
            {
                return this.conditionalFormatsField;
            }
            set
            {
                this.conditionalFormatsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 10)]
        public CT_ChartFormats chartFormats
        {
            get
            {
                return this.chartFormatsField;
            }
            set
            {
                this.chartFormatsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 11)]
        public CT_PivotHierarchies pivotHierarchies
        {
            get
            {
                return this.pivotHierarchiesField;
            }
            set
            {
                this.pivotHierarchiesField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 12)]
        public CT_PivotTableStyle pivotTableStyleInfo
        {
            get
            {
                return this.pivotTableStyleInfoField;
            }
            set
            {
                this.pivotTableStyleInfoField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 13)]
        public CT_PivotFilters filters
        {
            get
            {
                return this.filtersField;
            }
            set
            {
                this.filtersField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 14)]
        public CT_RowHierarchiesUsage rowHierarchiesUsage
        {
            get
            {
                return this.rowHierarchiesUsageField;
            }
            set
            {
                this.rowHierarchiesUsageField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 15)]
        public CT_ColHierarchiesUsage colHierarchiesUsage
        {
            get
            {
                return this.colHierarchiesUsageField;
            }
            set
            {
                this.colHierarchiesUsageField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 16)]
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint cacheId
        {
            get
            {
                return this.cacheIdField;
            }
            set
            {
                this.cacheIdField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool dataOnRows
        {
            get
            {
                return this.dataOnRowsField;
            }
            set
            {
                this.dataOnRowsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint dataPosition
        {
            get
            {
                return this.dataPositionField;
            }
            set
            {
                this.dataPositionField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool dataPositionSpecified
        {
            get
            {
                return this.dataPositionFieldSpecified;
            }
            set
            {
                this.dataPositionFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint autoFormatId
        {
            get
            {
                return this.autoFormatIdField;
            }
            set
            {
                this.autoFormatIdField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool autoFormatIdSpecified
        {
            get
            {
                return this.autoFormatIdFieldSpecified;
            }
            set
            {
                this.autoFormatIdFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyNumberFormats
        {
            get
            {
                return this.applyNumberFormatsField;
            }
            set
            {
                this.applyNumberFormatsField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyNumberFormatsSpecified
        {
            get
            {
                return this.applyNumberFormatsFieldSpecified;
            }
            set
            {
                this.applyNumberFormatsFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyBorderFormats
        {
            get
            {
                return this.applyBorderFormatsField;
            }
            set
            {
                this.applyBorderFormatsField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyBorderFormatsSpecified
        {
            get
            {
                return this.applyBorderFormatsFieldSpecified;
            }
            set
            {
                this.applyBorderFormatsFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyFontFormats
        {
            get
            {
                return this.applyFontFormatsField;
            }
            set
            {
                this.applyFontFormatsField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyFontFormatsSpecified
        {
            get
            {
                return this.applyFontFormatsFieldSpecified;
            }
            set
            {
                this.applyFontFormatsFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyPatternFormats
        {
            get
            {
                return this.applyPatternFormatsField;
            }
            set
            {
                this.applyPatternFormatsField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyPatternFormatsSpecified
        {
            get
            {
                return this.applyPatternFormatsFieldSpecified;
            }
            set
            {
                this.applyPatternFormatsFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyAlignmentFormats
        {
            get
            {
                return this.applyAlignmentFormatsField;
            }
            set
            {
                this.applyAlignmentFormatsField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyAlignmentFormatsSpecified
        {
            get
            {
                return this.applyAlignmentFormatsFieldSpecified;
            }
            set
            {
                this.applyAlignmentFormatsFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyWidthHeightFormats
        {
            get
            {
                return this.applyWidthHeightFormatsField;
            }
            set
            {
                this.applyWidthHeightFormatsField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyWidthHeightFormatsSpecified
        {
            get
            {
                return this.applyWidthHeightFormatsFieldSpecified;
            }
            set
            {
                this.applyWidthHeightFormatsFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string dataCaption
        {
            get
            {
                return this.dataCaptionField;
            }
            set
            {
                this.dataCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string grandTotalCaption
        {
            get
            {
                return this.grandTotalCaptionField;
            }
            set
            {
                this.grandTotalCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string errorCaption
        {
            get
            {
                return this.errorCaptionField;
            }
            set
            {
                this.errorCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showError
        {
            get
            {
                return this.showErrorField;
            }
            set
            {
                this.showErrorField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string missingCaption
        {
            get
            {
                return this.missingCaptionField;
            }
            set
            {
                this.missingCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showMissing
        {
            get
            {
                return this.showMissingField;
            }
            set
            {
                this.showMissingField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string pageStyle
        {
            get
            {
                return this.pageStyleField;
            }
            set
            {
                this.pageStyleField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string pivotTableStyle
        {
            get
            {
                return this.pivotTableStyleField;
            }
            set
            {
                this.pivotTableStyleField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string vacatedStyle
        {
            get
            {
                return this.vacatedStyleField;
            }
            set
            {
                this.vacatedStyleField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string tag
        {
            get
            {
                return this.tagField;
            }
            set
            {
                this.tagField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(byte), "0")]
        public byte updatedVersion
        {
            get
            {
                return this.updatedVersionField;
            }
            set
            {
                this.updatedVersionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(byte), "0")]
        public byte minRefreshableVersion
        {
            get
            {
                return this.minRefreshableVersionField;
            }
            set
            {
                this.minRefreshableVersionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool asteriskTotals
        {
            get
            {
                return this.asteriskTotalsField;
            }
            set
            {
                this.asteriskTotalsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showItems
        {
            get
            {
                return this.showItemsField;
            }
            set
            {
                this.showItemsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool editData
        {
            get
            {
                return this.editDataField;
            }
            set
            {
                this.editDataField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool disableFieldList
        {
            get
            {
                return this.disableFieldListField;
            }
            set
            {
                this.disableFieldListField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showCalcMbrs
        {
            get
            {
                return this.showCalcMbrsField;
            }
            set
            {
                this.showCalcMbrsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool visualTotals
        {
            get
            {
                return this.visualTotalsField;
            }
            set
            {
                this.visualTotalsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showMultipleLabel
        {
            get
            {
                return this.showMultipleLabelField;
            }
            set
            {
                this.showMultipleLabelField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showDataDropDown
        {
            get
            {
                return this.showDataDropDownField;
            }
            set
            {
                this.showDataDropDownField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showDrill
        {
            get
            {
                return this.showDrillField;
            }
            set
            {
                this.showDrillField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool printDrill
        {
            get
            {
                return this.printDrillField;
            }
            set
            {
                this.printDrillField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showMemberPropertyTips
        {
            get
            {
                return this.showMemberPropertyTipsField;
            }
            set
            {
                this.showMemberPropertyTipsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showDataTips
        {
            get
            {
                return this.showDataTipsField;
            }
            set
            {
                this.showDataTipsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool enableWizard
        {
            get
            {
                return this.enableWizardField;
            }
            set
            {
                this.enableWizardField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool enableDrill
        {
            get
            {
                return this.enableDrillField;
            }
            set
            {
                this.enableDrillField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool enableFieldProperties
        {
            get
            {
                return this.enableFieldPropertiesField;
            }
            set
            {
                this.enableFieldPropertiesField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool preserveFormatting
        {
            get
            {
                return this.preserveFormattingField;
            }
            set
            {
                this.preserveFormattingField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool useAutoFormatting
        {
            get
            {
                return this.useAutoFormattingField;
            }
            set
            {
                this.useAutoFormattingField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint pageWrap
        {
            get
            {
                return this.pageWrapField;
            }
            set
            {
                this.pageWrapField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool pageOverThenDown
        {
            get
            {
                return this.pageOverThenDownField;
            }
            set
            {
                this.pageOverThenDownField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool subtotalHiddenItems
        {
            get
            {
                return this.subtotalHiddenItemsField;
            }
            set
            {
                this.subtotalHiddenItemsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool rowGrandTotals
        {
            get
            {
                return this.rowGrandTotalsField;
            }
            set
            {
                this.rowGrandTotalsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool colGrandTotals
        {
            get
            {
                return this.colGrandTotalsField;
            }
            set
            {
                this.colGrandTotalsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool fieldPrintTitles
        {
            get
            {
                return this.fieldPrintTitlesField;
            }
            set
            {
                this.fieldPrintTitlesField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool itemPrintTitles
        {
            get
            {
                return this.itemPrintTitlesField;
            }
            set
            {
                this.itemPrintTitlesField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool mergeItem
        {
            get
            {
                return this.mergeItemField;
            }
            set
            {
                this.mergeItemField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showDropZones
        {
            get
            {
                return this.showDropZonesField;
            }
            set
            {
                this.showDropZonesField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(byte), "0")]
        public byte createdVersion
        {
            get
            {
                return this.createdVersionField;
            }
            set
            {
                this.createdVersionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "1")]
        public uint indent
        {
            get
            {
                return this.indentField;
            }
            set
            {
                this.indentField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showEmptyRow
        {
            get
            {
                return this.showEmptyRowField;
            }
            set
            {
                this.showEmptyRowField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showEmptyCol
        {
            get
            {
                return this.showEmptyColField;
            }
            set
            {
                this.showEmptyColField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showHeaders
        {
            get
            {
                return this.showHeadersField;
            }
            set
            {
                this.showHeadersField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool compact
        {
            get
            {
                return this.compactField;
            }
            set
            {
                this.compactField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool outline
        {
            get
            {
                return this.outlineField;
            }
            set
            {
                this.outlineField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool outlineData
        {
            get
            {
                return this.outlineDataField;
            }
            set
            {
                this.outlineDataField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool compactData
        {
            get
            {
                return this.compactDataField;
            }
            set
            {
                this.compactDataField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool published
        {
            get
            {
                return this.publishedField;
            }
            set
            {
                this.publishedField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool gridDropZones
        {
            get
            {
                return this.gridDropZonesField;
            }
            set
            {
                this.gridDropZonesField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool immersive
        {
            get
            {
                return this.immersiveField;
            }
            set
            {
                this.immersiveField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool multipleFieldFilters
        {
            get
            {
                return this.multipleFieldFiltersField;
            }
            set
            {
                this.multipleFieldFiltersField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint chartFormat
        {
            get
            {
                return this.chartFormatField;
            }
            set
            {
                this.chartFormatField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string rowHeaderCaption
        {
            get
            {
                return this.rowHeaderCaptionField;
            }
            set
            {
                this.rowHeaderCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string colHeaderCaption
        {
            get
            {
                return this.colHeaderCaptionField;
            }
            set
            {
                this.colHeaderCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool fieldListSortAscending
        {
            get
            {
                return this.fieldListSortAscendingField;
            }
            set
            {
                this.fieldListSortAscendingField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool mdxSubqueries
        {
            get
            {
                return this.mdxSubqueriesField;
            }
            set
            {
                this.mdxSubqueriesField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool customListSort
        {
            get
            {
                return this.customListSortField;
            }
            set
            {
                this.customListSortField = value;
            }
        }


        public CT_PivotTableStyle AddNewPivotTableStyleInfo()
        {
            this.pivotTableStyleInfoField = new CT_PivotTableStyle();
            return this.pivotTableStyleInfoField;
        }

        public CT_RowFields AddNewRowFields()
        {
            this.rowFieldsField = new CT_RowFields();
            return this.rowFieldsField;
        }

        public CT_ColFields AddNewColFields()
        {
            this.colFieldsField = new CT_ColFields();
            return this.colFieldsField;
        }

        public CT_DataFields AddNewDataFields()
        {
            this.dataFieldsField = new CT_DataFields();
            return this.dataFieldsField;
        }

        public CT_PageFields AddNewPageFields()
        {
            this.pageFieldsField = new CT_PageFields();
            return this.pageFieldsField;
        }

        public CT_PivotFields AddNewPivotFields()
        {
            this.pivotFieldsField = new CT_PivotFields();
            return this.pivotFieldsField;
        }

        public CT_Location AddNewLocation()
        {
            this.locationField = new CT_Location();
            return this.locationField;
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

        private string refField;

        private uint firstHeaderRowField;

        private uint firstDataRowField;

        private uint firstDataColField;

        private uint rowPageCountField;

        private uint colPageCountField;

        public CT_Location()
        {
            this.rowPageCountField = ((uint)(0));
            this.colPageCountField = ((uint)(0));
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string @ref
        {
            get
            {
                return this.refField;
            }
            set
            {
                this.refField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint firstHeaderRow
        {
            get
            {
                return this.firstHeaderRowField;
            }
            set
            {
                this.firstHeaderRowField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint firstDataRow
        {
            get
            {
                return this.firstDataRowField;
            }
            set
            {
                this.firstDataRowField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint firstDataCol
        {
            get
            {
                return this.firstDataColField;
            }
            set
            {
                this.firstDataColField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint rowPageCount
        {
            get
            {
                return this.rowPageCountField;
            }
            set
            {
                this.rowPageCountField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint colPageCount
        {
            get
            {
                return this.colPageCountField;
            }
            set
            {
                this.colPageCountField = value;
            }
        }
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

        private List<CT_PivotField> pivotFieldField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_PivotFields()
        {
            this.pivotFieldField = new List<CT_PivotField>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(pivotField), Order = 0)]
        public List<CT_PivotField> pivotField
        {
            get
            {
                return this.pivotFieldField;
            }
            set
            {
                this.pivotFieldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }

        public void SetPivotFieldArray(int columnIndex, CT_PivotField pivotField)
        {
            this.pivotFieldField[columnIndex] = pivotField;
        }

        public CT_PivotField AddNewPivotField()
        {
            if (this.pivotFieldField == null)
                this.pivotFieldField = new List<CT_PivotField>();
            CT_PivotField f = new CT_PivotField();
            this.pivotFieldField.Add(f);
            return f;
        }

        public uint SizeOfPivotFieldArray()
        {
            if (this.pivotFieldField == null)
                this.pivotFieldField = new List<CT_PivotField>();
            return (uint)this.pivotFieldField.Count;
        }

        public CT_PivotField GetPivotFieldArray(int columnIndex)
        {
            if (this.pivotFieldField == null)
                this.pivotFieldField = new List<CT_PivotField>();
            return this.pivotFieldField[columnIndex];
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

        private CT_Items itemsField;

        private CT_AutoSortScope autoSortScopeField;

        private CT_ExtensionList extLstField;

        private string nameField;

        private ST_Axis axisField;

        private bool axisFieldSpecified;

        private bool dataFieldField;

        private string subtotalCaptionField;

        private bool showDropDownsField;

        private bool hiddenLevelField;

        private string uniqueMemberPropertyField;

        private bool compactField;

        private bool allDrilledField;

        private uint numFmtIdField;

        private bool numFmtIdFieldSpecified;

        private bool outlineField;

        private bool subtotalTopField;

        private bool dragToRowField;

        private bool dragToColField;

        private bool multipleItemSelectionAllowedField;

        private bool dragToPageField;

        private bool dragToDataField;

        private bool dragOffField;

        private bool showAllField;

        private bool insertBlankRowField;

        private bool serverFieldField;

        private bool insertPageBreakField;

        private bool autoShowField;

        private bool topAutoShowField;

        private bool hideNewItemsField;

        private bool measureFilterField;

        private bool includeNewItemsInFilterField;

        private uint itemPageCountField;

        private ST_FieldSortType sortTypeField;

        private bool dataSourceSortField;

        private bool dataSourceSortFieldSpecified;

        private bool nonAutoSortDefaultField;

        private uint rankByField;

        private bool rankByFieldSpecified;

        private bool defaultSubtotalField;

        private bool sumSubtotalField;

        private bool countASubtotalField;

        private bool avgSubtotalField;

        private bool maxSubtotalField;

        private bool minSubtotalField;

        private bool productSubtotalField;

        private bool countSubtotalField;

        private bool stdDevSubtotalField;

        private bool stdDevPSubtotalField;

        private bool varSubtotalField;

        private bool varPSubtotalField;

        private bool showPropCellField;

        private bool showPropTipField;

        private bool showPropAsCaptionField;

        private bool defaultAttributeDrillStateField;

        public CT_PivotField()
        {
            this.extLstField = new CT_ExtensionList();
            this.autoSortScopeField = new CT_AutoSortScope();
            this.itemsField = new CT_Items();
            this.dataFieldField = false;
            this.showDropDownsField = true;
            this.hiddenLevelField = false;
            this.compactField = true;
            this.allDrilledField = false;
            this.outlineField = true;
            this.subtotalTopField = true;
            this.dragToRowField = true;
            this.dragToColField = true;
            this.multipleItemSelectionAllowedField = false;
            this.dragToPageField = true;
            this.dragToDataField = true;
            this.dragOffField = true;
            this.showAllField = true;
            this.insertBlankRowField = false;
            this.serverFieldField = false;
            this.insertPageBreakField = false;
            this.autoShowField = false;
            this.topAutoShowField = true;
            this.hideNewItemsField = false;
            this.measureFilterField = false;
            this.includeNewItemsInFilterField = false;
            this.itemPageCountField = ((uint)(10));
            this.sortTypeField = ST_FieldSortType.manual;
            this.nonAutoSortDefaultField = false;
            this.defaultSubtotalField = true;
            this.sumSubtotalField = false;
            this.countASubtotalField = false;
            this.avgSubtotalField = false;
            this.maxSubtotalField = false;
            this.minSubtotalField = false;
            this.productSubtotalField = false;
            this.countSubtotalField = false;
            this.stdDevSubtotalField = false;
            this.stdDevPSubtotalField = false;
            this.varSubtotalField = false;
            this.varPSubtotalField = false;
            this.showPropCellField = false;
            this.showPropTipField = false;
            this.showPropAsCaptionField = false;
            this.defaultAttributeDrillStateField = false;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_Items items
        {
            get
            {
                return this.itemsField;
            }
            set
            {
                this.itemsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_AutoSortScope autoSortScope
        {
            get
            {
                return this.autoSortScopeField;
            }
            set
            {
                this.autoSortScopeField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 2)]
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_Axis axis
        {
            get
            {
                return this.axisField;
            }
            set
            {
                this.axisField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool axisSpecified
        {
            get
            {
                return this.axisFieldSpecified;
            }
            set
            {
                this.axisFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool dataField
        {
            get
            {
                return this.dataFieldField;
            }
            set
            {
                this.dataFieldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string subtotalCaption
        {
            get
            {
                return this.subtotalCaptionField;
            }
            set
            {
                this.subtotalCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showDropDowns
        {
            get
            {
                return this.showDropDownsField;
            }
            set
            {
                this.showDropDownsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool hiddenLevel
        {
            get
            {
                return this.hiddenLevelField;
            }
            set
            {
                this.hiddenLevelField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string uniqueMemberProperty
        {
            get
            {
                return this.uniqueMemberPropertyField;
            }
            set
            {
                this.uniqueMemberPropertyField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool compact
        {
            get
            {
                return this.compactField;
            }
            set
            {
                this.compactField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool allDrilled
        {
            get
            {
                return this.allDrilledField;
            }
            set
            {
                this.allDrilledField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint numFmtId
        {
            get
            {
                return this.numFmtIdField;
            }
            set
            {
                this.numFmtIdField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool numFmtIdSpecified
        {
            get
            {
                return this.numFmtIdFieldSpecified;
            }
            set
            {
                this.numFmtIdFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool outline
        {
            get
            {
                return this.outlineField;
            }
            set
            {
                this.outlineField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool subtotalTop
        {
            get
            {
                return this.subtotalTopField;
            }
            set
            {
                this.subtotalTopField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToRow
        {
            get
            {
                return this.dragToRowField;
            }
            set
            {
                this.dragToRowField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToCol
        {
            get
            {
                return this.dragToColField;
            }
            set
            {
                this.dragToColField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool multipleItemSelectionAllowed
        {
            get
            {
                return this.multipleItemSelectionAllowedField;
            }
            set
            {
                this.multipleItemSelectionAllowedField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToPage
        {
            get
            {
                return this.dragToPageField;
            }
            set
            {
                this.dragToPageField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToData
        {
            get
            {
                return this.dragToDataField;
            }
            set
            {
                this.dragToDataField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragOff
        {
            get
            {
                return this.dragOffField;
            }
            set
            {
                this.dragOffField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showAll
        {
            get
            {
                return this.showAllField;
            }
            set
            {
                this.showAllField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool insertBlankRow
        {
            get
            {
                return this.insertBlankRowField;
            }
            set
            {
                this.insertBlankRowField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool serverField
        {
            get
            {
                return this.serverFieldField;
            }
            set
            {
                this.serverFieldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool insertPageBreak
        {
            get
            {
                return this.insertPageBreakField;
            }
            set
            {
                this.insertPageBreakField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool autoShow
        {
            get
            {
                return this.autoShowField;
            }
            set
            {
                this.autoShowField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool topAutoShow
        {
            get
            {
                return this.topAutoShowField;
            }
            set
            {
                this.topAutoShowField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool hideNewItems
        {
            get
            {
                return this.hideNewItemsField;
            }
            set
            {
                this.hideNewItemsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool measureFilter
        {
            get
            {
                return this.measureFilterField;
            }
            set
            {
                this.measureFilterField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool includeNewItemsInFilter
        {
            get
            {
                return this.includeNewItemsInFilterField;
            }
            set
            {
                this.includeNewItemsInFilterField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "10")]
        public uint itemPageCount
        {
            get
            {
                return this.itemPageCountField;
            }
            set
            {
                this.itemPageCountField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_FieldSortType.manual)]
        public ST_FieldSortType sortType
        {
            get
            {
                return this.sortTypeField;
            }
            set
            {
                this.sortTypeField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool dataSourceSort
        {
            get
            {
                return this.dataSourceSortField;
            }
            set
            {
                this.dataSourceSortField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool dataSourceSortSpecified
        {
            get
            {
                return this.dataSourceSortFieldSpecified;
            }
            set
            {
                this.dataSourceSortFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool nonAutoSortDefault
        {
            get
            {
                return this.nonAutoSortDefaultField;
            }
            set
            {
                this.nonAutoSortDefaultField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint rankBy
        {
            get
            {
                return this.rankByField;
            }
            set
            {
                this.rankByField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool rankBySpecified
        {
            get
            {
                return this.rankByFieldSpecified;
            }
            set
            {
                this.rankByFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool defaultSubtotal
        {
            get
            {
                return this.defaultSubtotalField;
            }
            set
            {
                this.defaultSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool sumSubtotal
        {
            get
            {
                return this.sumSubtotalField;
            }
            set
            {
                this.sumSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool countASubtotal
        {
            get
            {
                return this.countASubtotalField;
            }
            set
            {
                this.countASubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool avgSubtotal
        {
            get
            {
                return this.avgSubtotalField;
            }
            set
            {
                this.avgSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool maxSubtotal
        {
            get
            {
                return this.maxSubtotalField;
            }
            set
            {
                this.maxSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool minSubtotal
        {
            get
            {
                return this.minSubtotalField;
            }
            set
            {
                this.minSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool productSubtotal
        {
            get
            {
                return this.productSubtotalField;
            }
            set
            {
                this.productSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool countSubtotal
        {
            get
            {
                return this.countSubtotalField;
            }
            set
            {
                this.countSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool stdDevSubtotal
        {
            get
            {
                return this.stdDevSubtotalField;
            }
            set
            {
                this.stdDevSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool stdDevPSubtotal
        {
            get
            {
                return this.stdDevPSubtotalField;
            }
            set
            {
                this.stdDevPSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool varSubtotal
        {
            get
            {
                return this.varSubtotalField;
            }
            set
            {
                this.varSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool varPSubtotal
        {
            get
            {
                return this.varPSubtotalField;
            }
            set
            {
                this.varPSubtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showPropCell
        {
            get
            {
                return this.showPropCellField;
            }
            set
            {
                this.showPropCellField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showPropTip
        {
            get
            {
                return this.showPropTipField;
            }
            set
            {
                this.showPropTipField = value;
            }
        }

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
        public bool defaultAttributeDrillState
        {
            get
            {
                return this.defaultAttributeDrillStateField;
            }
            set
            {
                this.defaultAttributeDrillStateField = value;
            }
        }

        public CT_Items AddNewItems()
        {
            this.itemsField = new CT_Items();
            return this.itemsField;
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

        private List<CT_Item> itemField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_Items()
        {
            this.itemField = new List<CT_Item>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(item), Order = 0)]
        public List<CT_Item> item
        {
            get
            {
                return this.itemField;
            }
            set
            {
                this.itemField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }

        public CT_Item AddNewItem()
        {
            if (this.itemField == null)
                this.itemField = new List<CT_Item>();
            CT_Item i = new CT_Item();
            this.itemField.Add(i);
            return i;
        }

        public uint SizeOfItemArray()
        {
            if (this.itemField == null)
                this.itemField = new List<CT_Item>();
            return (uint)this.itemField.Count;
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

        private string nField;

        private ST_ItemType tField;

        private bool hField;

        private bool sField;

        private bool sdField;

        private bool fField;

        private bool mField;

        private bool cField;

        private uint xField;

        private bool xFieldSpecified;

        private bool dField;

        private bool eField;

        public CT_Item()
        {
            this.tField = ST_ItemType.data;
            this.hField = false;
            this.sField = false;
            this.sdField = true;
            this.fField = false;
            this.mField = false;
            this.cField = false;
            this.dField = false;
            this.eField = true;
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string n
        {
            get
            {
                return this.nField;
            }
            set
            {
                this.nField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_ItemType.data)]
        public ST_ItemType t
        {
            get
            {
                return this.tField;
            }
            set
            {
                this.tField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool h
        {
            get
            {
                return this.hField;
            }
            set
            {
                this.hField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool s
        {
            get
            {
                return this.sField;
            }
            set
            {
                this.sField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool sd
        {
            get
            {
                return this.sdField;
            }
            set
            {
                this.sdField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool f
        {
            get
            {
                return this.fField;
            }
            set
            {
                this.fField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool m
        {
            get
            {
                return this.mField;
            }
            set
            {
                this.mField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool c
        {
            get
            {
                return this.cField;
            }
            set
            {
                this.cField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint x
        {
            get
            {
                return this.xField;
            }
            set
            {
                this.xField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool xSpecified
        {
            get
            {
                return this.xFieldSpecified;
            }
            set
            {
                this.xFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool d
        {
            get
            {
                return this.dField;
            }
            set
            {
                this.dField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool e
        {
            get
            {
                return this.eField;
            }
            set
            {
                this.eField = value;
            }
        }
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

        private CT_PivotArea pivotAreaField;

        public CT_AutoSortScope()
        {
            this.pivotAreaField = new CT_PivotArea();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_PivotArea pivotArea
        {
            get
            {
                return this.pivotAreaField;
            }
            set
            {
                this.pivotAreaField = value;
            }
        }
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

        private List<CT_Field> fieldField;

        private uint countField;

        public CT_RowFields()
        {
            this.fieldField = new List<CT_Field>();
            this.countField = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(field), Order = 0)]
        public List<CT_Field> field
        {
            get
            {
                return this.fieldField;
            }
            set
            {
                this.fieldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        public CT_Field AddNewField()
        {
            CT_Field f = new CT_Field();
            this.fieldField.Add(f);
            return f;
        }

        public uint SizeOfFieldArray()
        {
            return (uint)this.fieldField.Count;
        }

        public List<CT_Field> GetFieldArray()
        {
            return this.fieldField;
        }

        public CT_Field GetFieldArray(int p)
        {
            return this.fieldField[p];
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

        private int xField;

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int x
        {
            get
            {
                return this.xField;
            }
            set
            {
                this.xField = value;
            }
        }
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

        private List<CT_I> iField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_rowItems()
        {
            this.iField = new List<CT_I>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(i), Order = 0)]
        public List<CT_I> i
        {
            get
            {
                return this.iField;
            }
            set
            {
                this.iField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
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

        private List<CT_X> xField;

        private ST_ItemType tField;

        private uint rField;

        private uint iField;

        public CT_I()
        {
            this.xField = new List<CT_X>();
            this.tField = ST_ItemType.data;
            this.rField = ((uint)(0));
            this.iField = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(x), Order = 0)]
        public List<CT_X> x
        {
            get
            {
                return this.xField;
            }
            set
            {
                this.xField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_ItemType.data)]
        public ST_ItemType t
        {
            get
            {
                return this.tField;
            }
            set
            {
                this.tField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint r
        {
            get
            {
                return this.rField;
            }
            set
            {
                this.rField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint i
        {
            get
            {
                return this.iField;
            }
            set
            {
                this.iField = value;
            }
        }
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

        private List<CT_Field> fieldField;

        private uint countField;

        public CT_ColFields()
        {
            this.fieldField = new List<CT_Field>();
            this.countField = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(field), Order = 0)]
        public List<CT_Field> field
        {
            get
            {
                return this.fieldField;
            }
            set
            {
                this.fieldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        public uint SizeOfFieldArray()
        {
            if (this.fieldField == null)
                this.fieldField = new List<CT_Field>();
            return (uint)this.fieldField.Count;
        }

        public CT_Field AddNewField()
        {
            if (this.fieldField == null)
                this.fieldField = new List<CT_Field>();
            CT_Field f = new CT_Field();
            this.fieldField.Add(f);
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

        private List<CT_I> iField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_colItems()
        {
            this.iField = new List<CT_I>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(i), Order = 0)]
        public List<CT_I> i
        {
            get
            {
                return this.iField;
            }
            set
            {
                this.iField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
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

        private List<CT_PageField> pageFieldField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_PageFields()
        {
            this.pageFieldField = new List<CT_PageField>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(pageField), Order = 0)]
        public List<CT_PageField> pageField
        {
            get
            {
                return this.pageFieldField;
            }
            set
            {
                this.pageFieldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }

        public CT_PageField AddNewPageField()
        {
            if (this.pageFieldField == null)
                this.pageFieldField = new List<CT_PageField>();
            CT_PageField f = new CT_PageField();
            this.pageFieldField.Add(f);
            return f;
        }

        public uint SizeOfPageFieldArray()
        {
            if (this.pageFieldField == null)
                this.pageFieldField = new List<CT_PageField>();
            return (uint)this.pageFieldField.Count;
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

        private CT_ExtensionList extLstField;

        private int fldField;

        private uint itemField;

        private bool itemFieldSpecified;

        private int hierField;

        private bool hierFieldSpecified;

        private string nameField;

        private string capField;

        public CT_PageField()
        {
            this.extLstField = new CT_ExtensionList();
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int fld
        {
            get
            {
                return this.fldField;
            }
            set
            {
                this.fldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint item
        {
            get
            {
                return this.itemField;
            }
            set
            {
                this.itemField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool itemSpecified
        {
            get
            {
                return this.itemFieldSpecified;
            }
            set
            {
                this.itemFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int hier
        {
            get
            {
                return this.hierField;
            }
            set
            {
                this.hierField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool hierSpecified
        {
            get
            {
                return this.hierFieldSpecified;
            }
            set
            {
                this.hierFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string cap
        {
            get
            {
                return this.capField;
            }
            set
            {
                this.capField = value;
            }
        }
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

        private List<CT_DataField> dataFieldField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_DataFields()
        {
            this.dataFieldField = new List<CT_DataField>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(dataField), Order = 0)]
        public List<CT_DataField> dataField
        {
            get
            {
                return this.dataFieldField;
            }
            set
            {
                this.dataFieldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }

        public CT_DataField AddNewDataField()
        {
            if (this.dataFieldField == null)
                this.dataFieldField = new List<CT_DataField>();
            CT_DataField f = new CT_DataField();
            this.dataFieldField.Add(f);
            return f;
        }

        public uint SizeOfDataFieldArray()
        {
            return (uint)this.dataFieldField.Count;
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

        private CT_ExtensionList extLstField;

        private string nameField;

        private uint fldField;

        private ST_DataConsolidateFunction subtotalField;

        private ST_ShowDataAs showDataAsField;

        private int baseFieldField;

        private uint baseItemField;

        private uint numFmtIdField;

        private bool numFmtIdFieldSpecified;

        public CT_DataField()
        {
            this.extLstField = new CT_ExtensionList();
            this.subtotalField = ST_DataConsolidateFunction.sum;
            this.showDataAsField = ST_ShowDataAs.normal;
            this.baseFieldField = -1;
            this.baseItemField = ((uint)(1048832));
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint fld
        {
            get
            {
                return this.fldField;
            }
            set
            {
                this.fldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_DataConsolidateFunction.sum)]
        public ST_DataConsolidateFunction subtotal
        {
            get
            {
                return this.subtotalField;
            }
            set
            {
                this.subtotalField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_ShowDataAs.normal)]
        public ST_ShowDataAs showDataAs
        {
            get
            {
                return this.showDataAsField;
            }
            set
            {
                this.showDataAsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(-1)]
        public int baseField
        {
            get
            {
                return this.baseFieldField;
            }
            set
            {
                this.baseFieldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "1048832")]
        public uint baseItem
        {
            get
            {
                return this.baseItemField;
            }
            set
            {
                this.baseItemField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint numFmtId
        {
            get
            {
                return this.numFmtIdField;
            }
            set
            {
                this.numFmtIdField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool numFmtIdSpecified
        {
            get
            {
                return this.numFmtIdFieldSpecified;
            }
            set
            {
                this.numFmtIdFieldSpecified = value;
            }
        }
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

        private List<CT_Format> formatField;

        private uint countField;

        public CT_Formats()
        {
            this.formatField = new List<CT_Format>();
            this.countField = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(format), Order = 0)]
        public List<CT_Format> format
        {
            get
            {
                return this.formatField;
            }
            set
            {
                this.formatField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }
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

        private CT_PivotArea pivotAreaField;

        private CT_ExtensionList extLstField;

        private ST_FormatAction actionField;

        private uint dxfIdField;

        private bool dxfIdFieldSpecified;

        public CT_Format()
        {
            this.extLstField = new CT_ExtensionList();
            this.pivotAreaField = new CT_PivotArea();
            this.actionField = ST_FormatAction.formatting;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_PivotArea pivotArea
        {
            get
            {
                return this.pivotAreaField;
            }
            set
            {
                this.pivotAreaField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_FormatAction.formatting)]
        public ST_FormatAction action
        {
            get
            {
                return this.actionField;
            }
            set
            {
                this.actionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint dxfId
        {
            get
            {
                return this.dxfIdField;
            }
            set
            {
                this.dxfIdField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool dxfIdSpecified
        {
            get
            {
                return this.dxfIdFieldSpecified;
            }
            set
            {
                this.dxfIdFieldSpecified = value;
            }
        }
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

        private List<CT_ConditionalFormat> conditionalFormatField;

        private uint countField;

        public CT_ConditionalFormats()
        {
            this.conditionalFormatField = new List<CT_ConditionalFormat>();
            this.countField = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(conditionalFormat), Order = 0)]
        public List<CT_ConditionalFormat> conditionalFormat
        {
            get
            {
                return this.conditionalFormatField;
            }
            set
            {
                this.conditionalFormatField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }
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

        private CT_PivotAreas pivotAreasField;

        private CT_ExtensionList extLstField;

        private ST_Scope scopeField;

        private ST_Type typeField;

        private uint priorityField;

        public CT_ConditionalFormat()
        {
            this.extLstField = new CT_ExtensionList();
            this.pivotAreasField = new CT_PivotAreas();
            this.scopeField = ST_Scope.selection;
            this.typeField = ST_Type.none;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_PivotAreas pivotAreas
        {
            get
            {
                return this.pivotAreasField;
            }
            set
            {
                this.pivotAreasField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_Scope.selection)]
        public ST_Scope scope
        {
            get
            {
                return this.scopeField;
            }
            set
            {
                this.scopeField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_Type.none)]
        public ST_Type type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint priority
        {
            get
            {
                return this.priorityField;
            }
            set
            {
                this.priorityField = value;
            }
        }
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

        private List<CT_PivotArea> pivotAreaField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_PivotAreas()
        {
            this.pivotAreaField = new List<CT_PivotArea>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(pivotArea), Order = 0)]
        public List<CT_PivotArea> pivotArea
        {
            get
            {
                return this.pivotAreaField;
            }
            set
            {
                this.pivotAreaField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
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

        private List<CT_ChartFormat> chartFormatField;

        private uint countField;

        public CT_ChartFormats()
        {
            this.chartFormatField = new List<CT_ChartFormat>();
            this.countField = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(chartFormat), Order = 0)]
        public List<CT_ChartFormat> chartFormat
        {
            get
            {
                return this.chartFormatField;
            }
            set
            {
                this.chartFormatField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }
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

        private CT_PivotArea pivotAreaField;

        private uint chartField;

        private uint formatField;

        private bool seriesField;

        public CT_ChartFormat()
        {
            this.pivotAreaField = new CT_PivotArea();
            this.seriesField = false;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_PivotArea pivotArea
        {
            get
            {
                return this.pivotAreaField;
            }
            set
            {
                this.pivotAreaField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint chart
        {
            get
            {
                return this.chartField;
            }
            set
            {
                this.chartField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint format
        {
            get
            {
                return this.formatField;
            }
            set
            {
                this.formatField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool series
        {
            get
            {
                return this.seriesField;
            }
            set
            {
                this.seriesField = value;
            }
        }
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

        private List<CT_PivotHierarchy> pivotHierarchyField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_PivotHierarchies()
        {
            this.pivotHierarchyField = new List<CT_PivotHierarchy>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(pivotHierarchy), Order = 0)]
        public List<CT_PivotHierarchy> pivotHierarchy
        {
            get
            {
                return this.pivotHierarchyField;
            }
            set
            {
                this.pivotHierarchyField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
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

        private CT_MemberProperties mpsField;

        private List<CT_Members> membersField;

        private CT_ExtensionList extLstField;

        private bool outlineField;

        private bool multipleItemSelectionAllowedField;

        private bool subtotalTopField;

        private bool showInFieldListField;

        private bool dragToRowField;

        private bool dragToColField;

        private bool dragToPageField;

        private bool dragToDataField;

        private bool dragOffField;

        private bool includeNewItemsInFilterField;

        private string captionField;

        public CT_PivotHierarchy()
        {
            this.extLstField = new CT_ExtensionList();
            this.membersField = new List<CT_Members>();
            this.mpsField = new CT_MemberProperties();
            this.outlineField = false;
            this.multipleItemSelectionAllowedField = false;
            this.subtotalTopField = false;
            this.showInFieldListField = true;
            this.dragToRowField = true;
            this.dragToColField = true;
            this.dragToPageField = true;
            this.dragToDataField = false;
            this.dragOffField = true;
            this.includeNewItemsInFilterField = false;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_MemberProperties mps
        {
            get
            {
                return this.mpsField;
            }
            set
            {
                this.mpsField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(members), Order = 1)]
        public List<CT_Members> members
        {
            get
            {
                return this.membersField;
            }
            set
            {
                this.membersField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 2)]
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool outline
        {
            get
            {
                return this.outlineField;
            }
            set
            {
                this.outlineField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool multipleItemSelectionAllowed
        {
            get
            {
                return this.multipleItemSelectionAllowedField;
            }
            set
            {
                this.multipleItemSelectionAllowedField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool subtotalTop
        {
            get
            {
                return this.subtotalTopField;
            }
            set
            {
                this.subtotalTopField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showInFieldList
        {
            get
            {
                return this.showInFieldListField;
            }
            set
            {
                this.showInFieldListField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToRow
        {
            get
            {
                return this.dragToRowField;
            }
            set
            {
                this.dragToRowField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToCol
        {
            get
            {
                return this.dragToColField;
            }
            set
            {
                this.dragToColField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragToPage
        {
            get
            {
                return this.dragToPageField;
            }
            set
            {
                this.dragToPageField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool dragToData
        {
            get
            {
                return this.dragToDataField;
            }
            set
            {
                this.dragToDataField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dragOff
        {
            get
            {
                return this.dragOffField;
            }
            set
            {
                this.dragOffField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool includeNewItemsInFilter
        {
            get
            {
                return this.includeNewItemsInFilterField;
            }
            set
            {
                this.includeNewItemsInFilterField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string caption
        {
            get
            {
                return this.captionField;
            }
            set
            {
                this.captionField = value;
            }
        }
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

        private List<CT_MemberProperty> mpField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_MemberProperties()
        {
            this.mpField = new List<CT_MemberProperty>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(mp), Order = 0)]
        public List<CT_MemberProperty> mp
        {
            get
            {
                return this.mpField;
            }
            set
            {
                this.mpField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
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

        private string nameField;

        private bool showCellField;

        private bool showTipField;

        private bool showAsCaptionField;

        private uint nameLenField;

        private bool nameLenFieldSpecified;

        private uint pPosField;

        private bool pPosFieldSpecified;

        private uint pLenField;

        private bool pLenFieldSpecified;

        private uint levelField;

        private bool levelFieldSpecified;

        private uint fieldField;

        public CT_MemberProperty()
        {
            this.showCellField = false;
            this.showTipField = false;
            this.showAsCaptionField = false;
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showCell
        {
            get
            {
                return this.showCellField;
            }
            set
            {
                this.showCellField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showTip
        {
            get
            {
                return this.showTipField;
            }
            set
            {
                this.showTipField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showAsCaption
        {
            get
            {
                return this.showAsCaptionField;
            }
            set
            {
                this.showAsCaptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint nameLen
        {
            get
            {
                return this.nameLenField;
            }
            set
            {
                this.nameLenField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool nameLenSpecified
        {
            get
            {
                return this.nameLenFieldSpecified;
            }
            set
            {
                this.nameLenFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint pPos
        {
            get
            {
                return this.pPosField;
            }
            set
            {
                this.pPosField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool pPosSpecified
        {
            get
            {
                return this.pPosFieldSpecified;
            }
            set
            {
                this.pPosFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint pLen
        {
            get
            {
                return this.pLenField;
            }
            set
            {
                this.pLenField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool pLenSpecified
        {
            get
            {
                return this.pLenFieldSpecified;
            }
            set
            {
                this.pLenFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint level
        {
            get
            {
                return this.levelField;
            }
            set
            {
                this.levelField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool levelSpecified
        {
            get
            {
                return this.levelFieldSpecified;
            }
            set
            {
                this.levelFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint field
        {
            get
            {
                return this.fieldField;
            }
            set
            {
                this.fieldField = value;
            }
        }
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

        private List<CT_Member> memberField;

        private uint countField;

        private bool countFieldSpecified;

        private uint levelField;

        private bool levelFieldSpecified;

        public CT_Members()
        {
            this.memberField = new List<CT_Member>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(member), Order = 0)]
        public List<CT_Member> member
        {
            get
            {
                return this.memberField;
            }
            set
            {
                this.memberField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint level
        {
            get
            {
                return this.levelField;
            }
            set
            {
                this.levelField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool levelSpecified
        {
            get
            {
                return this.levelFieldSpecified;
            }
            set
            {
                this.levelFieldSpecified = value;
            }
        }
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

        private string nameField;

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }
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

        private string nameField;

        private bool showRowHeadersField;

        private bool showRowHeadersFieldSpecified;

        private bool showColHeadersField;

        private bool showColHeadersFieldSpecified;

        private bool showRowStripesField;

        private bool showRowStripesFieldSpecified;

        private bool showColStripesField;

        private bool showColStripesFieldSpecified;

        private bool showLastColumnField;

        private bool showLastColumnFieldSpecified;

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool showRowHeaders
        {
            get
            {
                return this.showRowHeadersField;
            }
            set
            {
                this.showRowHeadersField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool showRowHeadersSpecified
        {
            get
            {
                return this.showRowHeadersFieldSpecified;
            }
            set
            {
                this.showRowHeadersFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool showColHeaders
        {
            get
            {
                return this.showColHeadersField;
            }
            set
            {
                this.showColHeadersField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool showColHeadersSpecified
        {
            get
            {
                return this.showColHeadersFieldSpecified;
            }
            set
            {
                this.showColHeadersFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool showRowStripes
        {
            get
            {
                return this.showRowStripesField;
            }
            set
            {
                this.showRowStripesField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool showRowStripesSpecified
        {
            get
            {
                return this.showRowStripesFieldSpecified;
            }
            set
            {
                this.showRowStripesFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool showColStripes
        {
            get
            {
                return this.showColStripesField;
            }
            set
            {
                this.showColStripesField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool showColStripesSpecified
        {
            get
            {
                return this.showColStripesFieldSpecified;
            }
            set
            {
                this.showColStripesFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool showLastColumn
        {
            get
            {
                return this.showLastColumnField;
            }
            set
            {
                this.showLastColumnField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool showLastColumnSpecified
        {
            get
            {
                return this.showLastColumnFieldSpecified;
            }
            set
            {
                this.showLastColumnFieldSpecified = value;
            }
        }
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

        private List<CT_PivotFilter> filterField;

        private uint countField;

        public CT_PivotFilters()
        {
            this.filterField = new List<CT_PivotFilter>();
            this.countField = ((uint)(0));
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(filter), Order = 0)]
        public List<CT_PivotFilter> filter
        {
            get
            {
                return this.filterField;
            }
            set
            {
                this.filterField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }
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

        private CT_AutoFilter autoFilterField;

        private CT_ExtensionList extLstField;

        private uint fldField;

        private uint mpFldField;

        private bool mpFldFieldSpecified;

        private ST_PivotFilterType typeField;

        private int evalOrderField;

        private uint idField;

        private uint iMeasureHierField;

        private bool iMeasureHierFieldSpecified;

        private uint iMeasureFldField;

        private bool iMeasureFldFieldSpecified;

        private string nameField;

        private string descriptionField;

        private string stringValue1Field;

        private string stringValue2Field;

        public CT_PivotFilter()
        {
            this.extLstField = new CT_ExtensionList();
            this.autoFilterField = new CT_AutoFilter();
            this.evalOrderField = 0;
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 0)]
        public CT_AutoFilter autoFilter
        {
            get
            {
                return this.autoFilterField;
            }
            set
            {
                this.autoFilterField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Order = 1)]
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint fld
        {
            get
            {
                return this.fldField;
            }
            set
            {
                this.fldField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint mpFld
        {
            get
            {
                return this.mpFldField;
            }
            set
            {
                this.mpFldField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool mpFldSpecified
        {
            get
            {
                return this.mpFldFieldSpecified;
            }
            set
            {
                this.mpFldFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_PivotFilterType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public int evalOrder
        {
            get
            {
                return this.evalOrderField;
            }
            set
            {
                this.evalOrderField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint iMeasureHier
        {
            get
            {
                return this.iMeasureHierField;
            }
            set
            {
                this.iMeasureHierField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool iMeasureHierSpecified
        {
            get
            {
                return this.iMeasureHierFieldSpecified;
            }
            set
            {
                this.iMeasureHierFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint iMeasureFld
        {
            get
            {
                return this.iMeasureFldField;
            }
            set
            {
                this.iMeasureFldField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool iMeasureFldSpecified
        {
            get
            {
                return this.iMeasureFldFieldSpecified;
            }
            set
            {
                this.iMeasureFldFieldSpecified = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string description
        {
            get
            {
                return this.descriptionField;
            }
            set
            {
                this.descriptionField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string stringValue1
        {
            get
            {
                return this.stringValue1Field;
            }
            set
            {
                this.stringValue1Field = value;
            }
        }

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

        private List<CT_HierarchyUsage> rowHierarchyUsageField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_RowHierarchiesUsage()
        {
            this.rowHierarchyUsageField = new List<CT_HierarchyUsage>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(rowHierarchyUsage), Order = 0)]
        public List<CT_HierarchyUsage> rowHierarchyUsage
        {
            get
            {
                return this.rowHierarchyUsageField;
            }
            set
            {
                this.rowHierarchyUsageField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
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

        private int hierarchyUsageField;

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int hierarchyUsage
        {
            get
            {
                return this.hierarchyUsageField;
            }
            set
            {
                this.hierarchyUsageField = value;
            }
        }
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

        private List<CT_HierarchyUsage> colHierarchyUsageField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_ColHierarchiesUsage()
        {
            this.colHierarchyUsageField = new List<CT_HierarchyUsage>();
        }

        [System.Xml.Serialization.XmlElementAttribute(nameof(colHierarchyUsage), Order = 0)]
        public List<CT_HierarchyUsage> colHierarchyUsage
        {
            get
            {
                return this.colHierarchyUsageField;
            }
            set
            {
                this.colHierarchyUsageField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
    }
    
}
