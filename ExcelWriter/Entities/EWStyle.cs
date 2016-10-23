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
            foreach (var item in styles)
            {
                if (item.Font == null)
                {
                    fontList.Add(GetDefaultFont());
                }
                else
                {
                    fontList.Add(item.Font.oxFont);
                }

                if (item.Fill == null)
                {
                    fillList.Add(GetDefaultFill());
                }
                else
                {
                    fillList.Add(item.Fill.oxFill);
                }

                if (item.Border == null)
                {
                    borderList.Add(new EWBorder().GetOpenXmlBorder());
                }
                else
                {
                    borderList.Add(item.Border.GetOpenXmlBorder());
                }

                cellFormatList.Add(new CellFormat() { FontId = fontId++, FillId = fillId++, BorderId = borderId++, ApplyFont = true, ApplyFill = true });

                selectors.Add(item.Selector, (fontId - 1).ToString());
            }

            return new Stylesheet(new Fonts(fontList),
                new Fills(fillList),
                new Borders(borderList),
                new CellFormats(cellFormatList));
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
