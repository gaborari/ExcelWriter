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

            _oxFont = new Font();
            FontSize fontSize1 = new FontSize() { Val = fontSize };
            _oxFont.Append(fontSize1);

            Color color1 = new Color() { Theme = color.ToUint() };
            _oxFont.Append(color1);

            FontName fontName1 = new FontName() { Val = fontName };
            _oxFont.Append(fontName1);

            if (isBold == true)
            {
                var bold1 = new Bold();
                _oxFont.Append(bold1);
            }

            if (isItalic == true)
            {
                var italic1 = new Italic();
                _oxFont.Append(italic1);
            }
            
        }
    }
}
