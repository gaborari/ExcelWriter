using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWriter.Entities
{
    public partial class EWStyle
    {
        public EWStyle()
        {
        }

        public static Font GetDefaultFont()
        {
            return new Font(                                                               // Index 0 – The default font.
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }
                        );
        }

        public static Fill GetDefaultFill()
        {
            return new Fill(                                                           // Index 0 – The default fill.
                        new PatternFill() { PatternType = PatternValues.None }
                    );
        }

        public static Border GetDefaultBorder()
        {
            return new Border(                                                         // Index 0 – The default border.
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder());
        }

        public static CellFormat GetDefaultCellFormat()
        {
            return new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 };
        }

        //TODO remove
        public static Stylesheet GetBaseStyle()
        {
            return  new Stylesheet(
                new Fonts(
                    GetDefaultFont(),
                    (new EWFont(11, "Calibri", "000000", isBold: true))._oxFont,
                    (new EWFont(11, "Calibri", "000000", isItalic: true))._oxFont,
                    (new EWFont(16, "Times New Roman", "000000"))._oxFont
                ),
                new Fills(
                    GetDefaultFill(),
                    (new EWFill(EWPattern.Gray125)).OxFill,
                    (new EWFill(EWPattern.Solid, foreGroundColor: "FFFFFF00")).OxFill            // Index 2 – The yellow fill.
                ),
                new Borders(
                    GetDefaultBorder(),
                    new Border(                                                         // Index 1 – Applies a Left, Right, Top, Bottom border to a cell
                        new LeftBorder(
                            new Color() { Auto = true }
                        ) { Style = BorderStyleValues.Thin },
                        new RightBorder(
                            new Color() { Auto = true }
                        ) { Style = BorderStyleValues.Thin },
                        new TopBorder(
                            new Color() { Auto = true }
                        ) { Style = BorderStyleValues.Thin },
                        new BottomBorder(
                            new Color() { Auto = true }
                        ) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                ),
                new CellFormats(
                    GetDefaultCellFormat(),                          // Index 0 – The default cell style.  If a cell does not have a style index applied it will use this style combination instead
                    new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 1 – Bold 
                    new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 2 – Italic
                    new CellFormat() { FontId = 3, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 3 – Times Roman
                    new CellFormat() { FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true },       // Index 4 – Yellow Fill
                    new CellFormat(                                                                   // Index 5 – Alignment
                        new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
                    ) { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }      // Index 6 – Border
                ));
        }
    }
}
