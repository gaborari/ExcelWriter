using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelWriter.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelWriter.Entities
{
    public partial class EWStyle
    {
        public EWBorder Border { get; set; }
        public EWFont Font { get; set; }
        public EWFill Fill { get; set; }
        public string Selector { get; set; }

        internal static Dictionary<string, string> selectors { get; set; } = new Dictionary<string, string>();

        internal static Font _defaultFont;
        internal static Font DefaultFont
        {
            get
            {
                if (_defaultFont == null)
                {
                    _defaultFont = GetDefaultFont();
                }
                return _defaultFont;
            }
        }

        internal static Fill _defaultFill;
        internal static Fill DefaultFill
        {
            get
            {
                if (_defaultFill == null)
                {
                    _defaultFill = GetDefaultFill();
                }
                return _defaultFill;
            }
        }


        internal static Border _defaultBorder;
        internal static Border DefaultBorder
        {
            get
            {
                if (_defaultBorder == null)
                {
                    _defaultBorder = new EWBorder().GetOpenXmlBorder();
                }
                return _defaultBorder;
            }
        }

        public EWStyle(string selector)
        {
            this.Selector = selector;
        }


        private static Font GetDefaultFont()
        {
            return new Font(                                                               // Index 0 – The default font.
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = System.Drawing.Color.Black.ToHexString() } },
                        new FontName() { Val = "Calibri" }
                        );
        }

        private static Fill GetDefaultFill()
        {
            return new Fill(                                                           // Index 0 – The default fill.
                        new PatternFill() { PatternType = PatternValues.None }
                    );
        }

        public static CellFormat GetDefaultCellFormat()
        {
            return new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 };
        }

        internal static Stylesheet GetStyleSheet(IEnumerable<EWStyle> styles)
        {
            //return default style;
            if (styles == null || !styles.Any())
            {
                return new Stylesheet(
                    new Fonts(DefaultFont),
                    new Fills(DefaultFill),
                    new Borders(DefaultBorder),
                    new CellFormats(GetDefaultCellFormat())
                    );
            }

            var fontList = new List<Font>();
            var fillList = new List<Fill>();
            var borderList = new List<Border>();
            var cellFormatList = new List<CellFormat>();

            UInt32 borderId, fillId, fontId;
            borderId = fillId = fontId = 0;
            //foreach (var item in styles)
            //{
            //    if (item.Font == null)
            //    {
            //        fontList.Add(GetDefaultFont());
            //    }
            //    else
            //    {
            //        fontList.Add(item.Font.oxFont);
            //    }

            //    if (item.Fill == null)
            //    {
            //        fillList.Add(GetDefaultFill());
            //    }
            //    else
            //    {
            //        fillList.Add(item.Fill.oxFill);
            //    }

            //    if (item.Border == null)
            //    {
            //        borderList.Add(new EWBorder().GetOpenXmlBorder());
            //    }
            //    else
            //    {
            //        borderList.Add(item.Border.GetOpenXmlBorder());
            //    }

            //    CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            //    CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };
            //    CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };
            //    CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };


            //    cellFormatList.Add(new CellFormat() { FontId = fontId++, FillId = fillId++, BorderId = borderId++, ApplyFill = true });

            //    selectors.Add(item.Selector, (fontId - 1).ToString());
            //}

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)1U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            fonts1.Append(font1);

            Fills fills1 = new Fills() { Count = (UInt32Value)5U };
            
            // FillId = 0
            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };
            fill1.Append(patternFill1);

            // FillId = 1
            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };
            fill2.Append(patternFill2);

            // FillId = 2,RED
            Fill fill3 = new Fill();
            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "FFFF0000" };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);
            fill3.Append(patternFill3);

            // FillId = 3,BLUE
            Fill fill4 = new Fill();
            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = "FF0070C0" };
            BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill4.Append(foregroundColor2);
            patternFill4.Append(backgroundColor2);
            fill4.Append(patternFill4);

            // FillId = 4,YELLO
            Fill fill5 = new Fill();
            PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor3 = new ForegroundColor() { Rgb = "FFFFFF00" };
            BackgroundColor backgroundColor3 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill5.Append(foregroundColor3);
            patternFill5.Append(backgroundColor3);
            fill5.Append(patternFill5);

            fills1.Append(fill1);
            fills1.Append(fill2);
            fills1.Append(fill3);
            fills1.Append(fill4);
            fills1.Append(fill5);

            Borders borders1 = new Borders() { Count = (UInt32Value)1U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            borders1.Append(border1);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)4U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };
            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);
            cellFormats1.Append(cellFormat5);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };


            return new Stylesheet(
                    fonts1,
                    fills1,
                    borders1,
                    cellFormats1
                );
        }

        ////TODO remove
        //public static Stylesheet GetBaseStyle()
        //{
        //    return new Stylesheet(
        //        new Fonts(
        //            GetDefaultFont(),
        //            (new EWFont(11, "Calibri", "000000", isBold: true))._oxFont,
        //            (new EWFont(11, "Calibri", "000000", isItalic: true))._oxFont,
        //            (new EWFont(16, "Times New Roman", "000000"))._oxFont
        //        ),
        //        new Fills(
        //            GetDefaultFill(),
        //            (new EWFill(EWPattern.Gray125)).oxFill,
        //            (new EWFill(EWPattern.Solid, foreGroundColor: "FFFFFF00")).oxFill            // Index 2 – The yellow fill.
        //        ),
        //        new Borders(
        //            new EWBorder().GetOpenXmlBorder(),
        //            new EWBorder().GetOpenXmlBorder(),
        //            new EWBorder().GetOpenXmlBorder()
        //        ),
        //        new CellFormats(
        //            GetDefaultCellFormat(),                          // Index 0 – The default cell style.  If a cell does not have a style index applied it will use this style combination instead
        //            new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 1 – Bold 
        //            new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 2 – Italic
        //            new CellFormat() { FontId = 3, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 3 – Times Roman
        //            new CellFormat() { FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true },       // Index 4 – Yellow Fill
        //            new CellFormat(                                                                   // Index 5 – Alignment
        //                new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
        //            )
        //            { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },
        //            new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }      // Index 6 – Border
        //        ));
        //}
    }
}
