using NPOI.OpenXml4Net.Util;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace NPOI.OpenXmlFormats.Spreadsheet
{
    public class MapInfoDocument
    {
        CT_MapInfo map = null;

        public MapInfoDocument()
        {
        }
        public MapInfoDocument(CT_MapInfo map)
        {
            this.map = map;
        }
        public CT_MapInfo GetMapInfo()
        {
            return this.map;
        }
        public void SetMapInfo(CT_MapInfo map)
        {
            this.map = map;
        }
        public void SetComments(CT_MapInfo map)
        {
            this.map = map;
        }
        public void Save(Stream stream)
        {
            using (var sw = new StreamWriter(stream))
            {
                sw.Write("<MapInfo xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" SelectionNamespaces=\"xmlns:ns1='http://schemas.openxmlformats.org/spreadsheetml/2006/main'\">");
                if (this.map.Schema != null)
                {
                    foreach (CT_Schema ctSchema in this.map.Schema)
                    {
                        sw.Write($"<Schema ID=\"{ctSchema.ID}\"");
                        if (ctSchema.Namespace != null)
                            sw.Write($" Namespace=\"{ctSchema.Namespace}\"");
                        if (ctSchema.SchemaRef != null)
                            sw.Write($" SchemaRef=\"{ctSchema.SchemaRef}\"");
                        sw.Write(">");
                        sw.Write(ctSchema.InnerXml);
                        sw.Write("</Schema>");
                    }
                }
                if (this.map.Map != null)
                {
                    foreach (CT_Map ctMap in this.map.Map)
                    {
                        sw.Write($"<Map ID=\"{ctMap.ID}\"");
                        if (ctMap.SchemaID != null)
                            sw.Write($" SchemaID=\"{ctMap.SchemaID}\"");
                        if (ctMap.RootElement != null)
                            sw.Write($" RootElement=\"{ctMap.RootElement}\"");
                        if (ctMap.Name != null)
                            sw.Write($" Name=\"{ctMap.Name}\"");
                        if (ctMap.PreserveFormat)
                            sw.Write(" PreserveFormat=\"true\"");
                        if (ctMap.PreserveSortAFLayout)
                            sw.Write(" PreserveSortAFLayout=\"true\"");
                        if (ctMap.ShowImportExportValidationErrors)
                            sw.Write(" ShowImportExportValidationErrors=\"true\"");
                        if (ctMap.Append)
                            sw.Write(" Append=\"true\"");
                        if (ctMap.AutoFit)
                            sw.Write(" AutoFit=\"true\"");
                        sw.Write(" />");
                    }
                }
                sw.Write("</MapInfo");
            }
        }

        public static MapInfoDocument Parse(XmlDocument xmldoc, XmlNamespaceManager NameSpaceManager)
        {
            var doc = new MapInfoDocument
            {
                map = new CT_MapInfo()
            };
            doc.map.Map = new System.Collections.Generic.List<CT_Map>();
            foreach (XmlElement mapNode in xmldoc.SelectNodes("d:MapInfo/d:Map", NameSpaceManager))
            {
                var ctMap = new CT_Map
                {
                    ID = XmlHelper.ReadUInt(mapNode.GetAttributeNode(nameof(CT_Map.ID))),
                    Name = XmlHelper.ReadString(mapNode.GetAttributeNode(nameof(CT_Map.Name))),
                    RootElement = XmlHelper.ReadString(mapNode.GetAttributeNode(nameof(CT_Map.RootElement))),
                    SchemaID = XmlHelper.ReadString(mapNode.GetAttributeNode(nameof(CT_Map.SchemaID))),
                    ShowImportExportValidationErrors = XmlHelper.ReadBool(mapNode.GetAttributeNode(nameof(CT_Map.ShowImportExportValidationErrors))),
                    PreserveFormat = XmlHelper.ReadBool(mapNode.GetAttributeNode(nameof(CT_Map.PreserveFormat))),
                    PreserveSortAFLayout = XmlHelper.ReadBool(mapNode.GetAttributeNode(nameof(CT_Map.PreserveSortAFLayout))),
                    Append = XmlHelper.ReadBool(mapNode.GetAttributeNode(nameof(CT_Map.Append))),
                    AutoFit = XmlHelper.ReadBool(mapNode.GetAttributeNode(nameof(CT_Map.AutoFit)))
                };
                doc.map.Map.Add(ctMap);
            }
            doc.map.Schema = new System.Collections.Generic.List<CT_Schema>();
            foreach (XmlNode schemaNode in xmldoc.SelectNodes("d:MapInfo/d:Schema", NameSpaceManager))
            {
                var ctSchema = new CT_Schema();
                ctSchema.ID = schemaNode.Attributes[nameof(CT_Schema.ID)].Value;
                if (schemaNode.Attributes[nameof(CT_Schema.Namespace)] != null)
                    ctSchema.Namespace = schemaNode.Attributes[nameof(CT_Schema.Namespace)].Value;
                if (schemaNode.Attributes[nameof(CT_Schema.SchemaRef)] != null)
                    ctSchema.SchemaRef = schemaNode.Attributes[nameof(CT_Schema.SchemaRef)].Value;
                ctSchema.InnerXml = schemaNode.InnerXml;
                doc.map.Schema.Add(ctSchema);
            }
            return doc;
        }
    }
}
