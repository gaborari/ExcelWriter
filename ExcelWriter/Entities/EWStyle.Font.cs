using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelWriter.Helpers;

namespace ExcelWriter.Entities
{
    public class EWFont
    {
        internal Font _oxFont;

        internal Font oxFont
        {
            get
            {
                return _oxFont;
            }
        }

        public EWFont(double fontSize, string fontName, System.Drawing.Color color, bool isBold = false, bool isItalic = false)
        {
            var fontColor = color.ToHexString();

            if (isBold && isItalic)
            {
                _oxFont = new Font(
                                new Bold(),
                                new Italic(),
                                new FontSize() { Val = fontSize },
                                new Color() { Rgb = new HexBinaryValue() { Value = fontColor } },
                                new FontName() { Val = fontName }
                            );

                return;
            }

            if (isBold && !isItalic)
            {
                _oxFont = new Font(
                                new Bold(),
                                new FontSize() { Val = fontSize },
                                new Color() { Rgb = new HexBinaryValue() { Value = fontColor } },
                                new FontName() { Val = fontName }
                            );

                return;
            }

            if (!isBold && isItalic)
            {
                _oxFont = new Font(
                                new Italic(),
                                new FontSize() { Val = fontSize },
                                new Color() { Rgb = new HexBinaryValue() { Value = fontColor } },
                                new FontName() { Val = fontName }
                            );

                return;
            }

            _oxFont = new Font
            (
                new FontSize() { Val = fontSize },
                new Color() { Rgb = new HexBinaryValue() { Value = fontColor } },
                new FontName() { Val = fontName }
            );
            
        }
    }
}
