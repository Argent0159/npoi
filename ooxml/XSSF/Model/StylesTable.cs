/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using NPOI.OpenXml4Net.OPC;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.UserModel.Extensions;
using System.Collections.ObjectModel;
using NPOI.SS;

namespace NPOI.XSSF.Model
{


    /**
     * Table of styles shared across all sheets in a workbook.
     *
     * @author ugo
     */
    public class StylesTable : POIXMLDocumentPart
    {
        private Dictionary<int, string> numberFormats = new Dictionary<int, string>();
        private bool[] usedNumberFormats = new bool[SpreadsheetVersion.EXCEL2007.MaxCellStyles];
        private List<XSSFFont> fonts = new List<XSSFFont>();
        private List<XSSFCellFill> fills = new List<XSSFCellFill>();
        private List<XSSFCellBorder> borders = new List<XSSFCellBorder>();
        private List<CT_Xf> styleXfs = new List<CT_Xf>();
        private List<CT_Xf> xfs = new List<CT_Xf>();

        private List<CT_Dxf> dxfs = new List<CT_Dxf>();

        /**
         * The first style id available for use as a custom style
         */
        public static int FIRST_CUSTOM_STYLE_ID = BuiltinFormats.FIRST_USER_DEFINED_FORMAT_INDEX + 1;
        private static int MAXIMUM_STYLE_ID = SpreadsheetVersion.EXCEL2007.MaxCellStyles;

        private StyleSheetDocument doc;
        private ThemesTable theme;

        /**
         * Create a new, empty StylesTable
         */
        public StylesTable()
            : base()
        {

            doc = new StyleSheetDocument();
            doc.AddNewStyleSheet();
            // Initialization required in order to make the document Readable by MSExcel
            Initialize();
        }

        internal StylesTable(PackagePart part, PackageRelationship rel)
            : base(part, rel)
        {
            XmlDocument xmldoc = ConvertStreamToXml(part.GetInputStream());
            ReadFrom(xmldoc);
        }

        public ThemesTable GetTheme()
        {
            return theme;
        }

        public void SetTheme(ThemesTable theme)
        {
            this.theme = theme;

            // Pass the themes table along to things which need to 
            //  know about it, but have already been Created by now
            foreach (XSSFFont font in fonts)
            {
                font.SetThemesTable(theme);
            }
            foreach (XSSFCellBorder border in borders)
            {
                border.SetThemesTable(theme);
            }
        }

        /**
         * Read this shared styles table from an XML file.
         *
         * @param is The input stream Containing the XML document.
         * @throws IOException if an error occurs while Reading.
         */

        protected void ReadFrom(XmlDocument xmldoc)
        {
            try
            {
                
                doc = StyleSheetDocument.Parse(xmldoc, NamespaceManager);

                CT_Stylesheet styleSheet = doc.GetStyleSheet();

                // Grab all the different bits we care about
                CT_NumFmts ctfmts = styleSheet.numFmts;
                if (ctfmts != null)
                {
                    foreach (CT_NumFmt nfmt in ctfmts.numFmt)
                    {
                        int formatId = (int)nfmt.numFmtId;
                        numberFormats.Add(formatId, nfmt.formatCode);
                        usedNumberFormats[formatId] = true;
                    }
                }

                CT_Fonts ctfonts = styleSheet.fonts;
                if (ctfonts != null)
                {
                    int index = 0;
                    foreach (CT_Font font in ctfonts.font)
                    {
                        // Create the font and save it. Themes Table supplied later
                        XSSFFont f = new XSSFFont(font, index);
                        fonts.Add(f);
                        index++;
                    }
                }
                CT_Fills ctFills = styleSheet.fills;
                if (ctFills != null)
                {
                    foreach (CT_Fill fill in ctFills.fill)
                    {
                        fills.Add(new XSSFCellFill(fill));
                    }
                }

                CT_Borders ctborders = styleSheet.borders;
                if (ctborders != null)
                {
                    foreach (CT_Border border in ctborders.border)
                    {
                        borders.Add(new XSSFCellBorder(border));
                    }
                }

                CT_CellXfs cellXfs = styleSheet.cellXfs;
                if (cellXfs != null)
                    xfs.AddRange(cellXfs.xf);

                CT_CellStyleXfs cellStyleXfs = styleSheet.cellStyleXfs;
                if (cellStyleXfs != null)
                    styleXfs.AddRange(cellStyleXfs.xf);

                CT_Dxfs styleDxfs = styleSheet.dxfs;
                if (styleDxfs != null)
                    dxfs.AddRange(styleDxfs.dxf);

            }
            catch (XmlException e)
            {
                throw new IOException(e.Message);
            }
        }

        // ===========================================================
        //  Start of style related Getters and Setters
        // ===========================================================

        public string GetNumberFormatAt(int index)
        {
            if (numberFormats.ContainsKey(index))
                return numberFormats[index];
            else
                return null;
        }

        public int PutNumberFormat(string fmt)
        {
            if (numberFormats.ContainsValue(fmt))
            {
                // Find the key, and return that
                foreach (KeyValuePair<int, string> numFmt in numberFormats)
                {
                    if (numFmt.Value.Equals(fmt))
                    {
                        return numFmt.Key;
                    }
                }
                throw new InvalidOperationException("Found the format, but couldn't figure out where - should never happen!");
            }

            // Find a spare key, and add that
            for (int i = FIRST_CUSTOM_STYLE_ID; i < usedNumberFormats.Length; i++)
            {
                if (!usedNumberFormats[i])
                {
                    usedNumberFormats[i] = true;
                    numberFormats.Add(i, fmt);
                    return i;
                }
            }

            throw new InvalidOperationException("The maximum number of Data Formats was exceeded. " +
              "You can define up to " + usedNumberFormats.Length + " formats in a .xlsx Workbook");
        }

        public XSSFFont GetFontAt(int index)
        {
            return fonts[index];
        }

        /**
         * Records the given font in the font table.
         * Will re-use an existing font index if this
         *  font matches another, EXCEPT if forced
         *  registration is requested.
         * This allows people to create several fonts
         *  then customise them later.
         * Note - End Users probably want to call
         *  {@link XSSFFont#registerTo(StylesTable)}
         */
        public int PutFont(XSSFFont font, bool forceRegistration)
        {
            int index = -1;
            if (!forceRegistration)
            {
                index = fonts.IndexOf(font);
            }

            if (index != -1)
            {
                return index;
            }

            index = fonts.Count;
            fonts.Add(font);
            return index;
        }
        public int PutFont(XSSFFont font)
        {
            return PutFont(font, false);
        }

        public XSSFCellStyle GetStyleAt(int index)
        {
            int styleXfId = 0;

            // 0 is the empty default
            if (xfs[index].xfId > 0)
            {
                styleXfId = (int)xfs[index].xfId;
            }

            return new XSSFCellStyle(index, styleXfId, this, theme);
        }
        public int PutStyle(XSSFCellStyle style)
        {
            CT_Xf mainXF = style.GetCoreXf();

            if (!xfs.Contains(mainXF))
            {
                xfs.Add(mainXF);
            }
            return xfs.IndexOf(mainXF);
        }

        public XSSFCellBorder GetBorderAt(int index)
        {
            return borders[index];
        }

        public int PutBorder(XSSFCellBorder border)
        {
            int index = borders.IndexOf(border);
            if (index != -1)
            {
                return index;
            }
            borders.Add(border);
            border.SetThemesTable(theme);
            return borders.Count - 1;
        }

        public XSSFCellFill GetFillAt(int index)
        {
            return fills[index];
        }

        public List<XSSFCellBorder> GetBorders()
        {
            return borders;
        }

        public ReadOnlyCollection<XSSFCellFill> GetFills()
        {
            return fills.AsReadOnly();
        }

        public List<XSSFFont> GetFonts()
        {
            return fonts;
        }

        public Dictionary<int, string> GetNumberFormats()
        {
            return numberFormats;
        }

        public int PutFill(XSSFCellFill fill)
        {
            int index = fills.IndexOf(fill);
            if (index != -1)
            {
                return index;
            }
            fills.Add(fill);
            return fills.Count - 1;
        }

        internal CT_Xf GetCellXfAt(int index)
        {
            return xfs[index];
        }
        internal int PutCellXf(CT_Xf cellXf)
        {
            xfs.Add(cellXf);
            return xfs.Count;
        }
        internal void ReplaceCellXfAt(int index, CT_Xf cellXf)
        {
            xfs[index] = cellXf;
        }

        internal CT_Xf GetCellStyleXfAt(int index)
        {
            if (index < 0 || index > styleXfs.Count)
                return null;
            return styleXfs[index];
        }
        internal int PutCellStyleXf(CT_Xf cellStyleXf)
        {
            styleXfs.Add(cellStyleXf);
            return styleXfs.Count;
        }
        internal void ReplaceCellStyleXfAt(int index, CT_Xf cellStyleXf)
        {
            styleXfs[index] = cellStyleXf;
        }

        /**
         * get the size of cell styles
         */
        public int NumCellStyles
        {
            get
            {
                // Each cell style has a unique xfs entry
                // Several might share the same styleXfs entry
                return xfs.Count;
            }
        }

        /**
         * For unit testing only
         */
        internal int NumberFormatSize
        {
            get
            {
                return numberFormats.Count;
            }
        }

        /**
         * For unit testing only
         */
        internal int XfsSize
        {
            get
            {
                return xfs.Count;
            }
        }
        /**
         * For unit testing only
         */
        internal int StyleXfsSize
        {
            get
            {
                return styleXfs.Count;
            }
        }
        /**
         * For unit testing only!
         */
        internal CT_Stylesheet GetCTStylesheet()
        {
            return doc.GetStyleSheet();
        }
        internal int DXfsSize
        {
            get
            {
                return dxfs.Count;
            }
        }


        /**
         * Write this table out as XML.
         *
         * @param out The stream to write to.
         * @throws IOException if an error occurs while writing.
         */
        public void WriteTo(Stream out1)
        {

            // Work on the current one
            // Need to do this, as we don't handle
            //  all the possible entries yet
            CT_Stylesheet styleSheet = doc.GetStyleSheet();

            // Formats
            CT_NumFmts ctFormats = new CT_NumFmts();
            ctFormats.count = (uint)numberFormats.Count;
            if (ctFormats.count > 0)
                ctFormats.countSpecified = true;

            for (int fmtId = 0; fmtId < usedNumberFormats.Length; fmtId++)
            {
                if (usedNumberFormats[fmtId])
                {
                    CT_NumFmt ctFmt = ctFormats.AddNewNumFmt();
                    ctFmt.numFmtId = (uint)(fmtId);
                    ctFmt.formatCode = (numberFormats[(fmtId)]);
                }
            }

            if (ctFormats.count>0)
                styleSheet.numFmts = ctFormats;

            // Fonts
            CT_Fonts ctFonts = styleSheet.fonts;
            if (ctFonts == null)
                ctFonts = new CT_Fonts();
            ctFonts.count = (uint)fonts.Count;
            if (ctFonts.count > 0)
                ctFonts.countSpecified = true;
            List<CT_Font> ctfnt = new List<CT_Font>(fonts.Count);

            foreach (XSSFFont f in fonts)
                ctfnt.Add(f.GetCTFont());
            ctFonts.SetFontArray(ctfnt);
            styleSheet.fonts = (ctFonts);

            // Fills
            CT_Fills ctFills = styleSheet.fills;
            if (ctFills == null)
            {
                ctFills = new CT_Fills();
            }
            ctFills.count = (uint)fills.Count;
            List<CT_Fill> ctf = new List<CT_Fill>(fills.Count);
            
            foreach (XSSFCellFill f in fills)
                ctf.Add( f.GetCTFill());
            ctFills.SetFillArray(ctf);
            if (ctFills.count > 0)
                ctFills.countSpecified = true;
            styleSheet.fills = ctFills;

            // Borders
            CT_Borders ctBorders = styleSheet.borders;
            if (ctBorders == null)
            {
                ctBorders = new CT_Borders();
            }
            ctBorders.count = (uint)borders.Count;
            List<CT_Border> ctb = new List<CT_Border>(borders.Count);
            foreach (XSSFCellBorder b in borders) 
                ctb.Add(b.GetCTBorder());
            
            ctBorders.SetBorderArray(ctb);
            styleSheet.borders = ctBorders;

            // Xfs
            if (xfs.Count > 0)
            {
                CT_CellXfs ctXfs = styleSheet.cellXfs;
                if (ctXfs == null)
                {
                    ctXfs = new CT_CellXfs();
                }
                ctXfs.count = (uint)xfs.Count;
                if (ctXfs.count > 0)
                    ctXfs.countSpecified = true;
                ctXfs.xf = xfs;

                styleSheet.cellXfs = (ctXfs);
            }

            // Style xfs
            if (styleXfs.Count > 0)
            {
                CT_CellStyleXfs ctSXfs = styleSheet.cellStyleXfs;
                if (ctSXfs == null)
                {
                    ctSXfs = new CT_CellStyleXfs();
                }
                ctSXfs.count = (uint)(styleXfs.Count);
                if (ctSXfs.count > 0)
                    ctSXfs.countSpecified = true;
                ctSXfs.xf = styleXfs;

                styleSheet.cellStyleXfs = (ctSXfs);
            }

            // Style dxfs
            if (dxfs.Count > 0)
            {
                CT_Dxfs ctDxfs = styleSheet.dxfs;
                if (ctDxfs == null)
                {
                    ctDxfs = new CT_Dxfs();
                }
                ctDxfs.count = (uint)dxfs.Count;
                if (ctDxfs.count > 0)
                    ctDxfs.countSpecified = true;
                ctDxfs.dxf = dxfs;

                styleSheet.dxfs = (ctDxfs);
            }

            // Save
            doc.Save(out1);
        }


        protected internal override void Commit()
        {
            PackagePart part = GetPackagePart();
            Stream out1 = part.GetOutputStream();
            WriteTo(out1);
            out1.Close();
        }

        private void Initialize()
        {
            //CT_Font ctFont = CreateDefaultFont();
            XSSFFont xssfFont = CreateDefaultFont();
            fonts.Add(xssfFont);

            CT_Fill[] ctFill = CreateDefaultFills();
            fills.Add(new XSSFCellFill(ctFill[0]));
            fills.Add(new XSSFCellFill(ctFill[1]));

            CT_Border ctBorder = CreateDefaultBorder();
            borders.Add(new XSSFCellBorder(ctBorder));

            CT_Xf styleXf = CreateDefaultXf();
            styleXfs.Add(styleXf);
            CT_Xf xf = CreateDefaultXf();
            xf.xfId = 0;
            xfs.Add(xf);
        }

        private static CT_Xf CreateDefaultXf()
        {
            CT_Xf ctXf = new CT_Xf();
            ctXf.numFmtId = 0;
            ctXf.fontId = 0;
            ctXf.fillId = 0;
            ctXf.borderId = 0;
            return ctXf;
        }
        private static CT_Border CreateDefaultBorder()
        {
            CT_Border ctBorder = new CT_Border();
            ctBorder.AddNewLeft();
            ctBorder.AddNewRight();
            ctBorder.AddNewTop();
            ctBorder.AddNewBottom();
            ctBorder.AddNewDiagonal();
            return ctBorder;
        }


        private static CT_Fill[] CreateDefaultFills()
        {
            CT_Fill[] ctFill = new CT_Fill[] { new CT_Fill(), new CT_Fill() };
            ctFill[0].AddNewPatternFill().patternType = (ST_PatternType.none);
            ctFill[1].AddNewPatternFill().patternType = (ST_PatternType.darkGray);
            return ctFill;
        }

        private static XSSFFont CreateDefaultFont()
        {
            CT_Font ctFont = new CT_Font();
            XSSFFont xssfFont = new XSSFFont(ctFont, 0);
            xssfFont.FontHeightInPoints = (XSSFFont.DEFAULT_FONT_SIZE);
            xssfFont.Color = (XSSFFont.DEFAULT_FONT_COLOR);//SetTheme
            xssfFont.FontName = (XSSFFont.DEFAULT_FONT_NAME);
            xssfFont.SetFamily(FontFamily.SWISS);
            xssfFont.SetScheme(FontScheme.MINOR);
            return xssfFont;
        }

        public CT_Dxf GetDxfAt(int index)
        {
            return dxfs[index];
        }

        public int PutDxf(CT_Dxf dxf)
        {
            this.dxfs.Add(dxf);
            return this.dxfs.Count;
        }

        public XSSFCellStyle CreateCellStyle()
        {
            int xfSize = styleXfs.Count;
            if (xfSize > MAXIMUM_STYLE_ID)
                throw new InvalidOperationException("The maximum number of Cell Styles was exceeded. " +
                          "You can define up to " + MAXIMUM_STYLE_ID + " style in a .xlsx Workbook");
        
            CT_Xf ctXf = new CT_Xf();
            ctXf.numFmtId = 0;
            ctXf.fontId = 0;
            ctXf.fillId = 0;
            ctXf.borderId = 0;
            ctXf.xfId = 0;
            
            int indexXf = PutCellXf(ctXf);
            return new XSSFCellStyle(indexXf - 1, xfSize - 1, this, theme);
        }

        /**
         * Finds a font that matches the one with the supplied attributes
         */
        public XSSFFont FindFont(short boldWeight, short color, short fontHeight, string name, bool italic, bool strikeout, FontSuperScript typeOffset,FontUnderlineType underline)
        {
            foreach (XSSFFont font in fonts)
            {
                if ((font.Boldweight == boldWeight)
                        && font.Color == color
                        && font.FontHeight == fontHeight
                        && font.FontName.Equals(name)
                        && font.IsItalic == italic
                        && font.IsStrikeout == strikeout
                        && font.TypeOffset == typeOffset
                        && font.Underline == underline)
                {
                    return font;
                }
            }
            return null;
        }
    }
}





