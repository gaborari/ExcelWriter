using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelWriter.Helpers;

namespace ExcelWriter.Entities
{
    public class EWFill
    {
        Fill _oxFill;
        public Fill oxFill
        {
            get
            {
                return _oxFill;
            }
        }

        public EWFill(System.Drawing.Color color, EWPattern pattern = EWPattern.Solid)
        {
            _oxFill = new Fill();

            PatternFill patternFill3 = new PatternFill() { PatternType = pattern.As<PatternValues>() };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = color.ToHexString() };
            //BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill3.Append(foregroundColor1);
            //patternFill3.Append(backgroundColor1);
            _oxFill.Append(patternFill3);

        }
    }

    public enum EWPattern
    {
        [EnumString("none")]
        None = 0,
        [EnumString("solid")]
        Solid = 1,
        [EnumString("mediumGray")]
        MediumGray = 2,
        [EnumString("darkGray")]
        DarkGray = 3,
        [EnumString("lightGray")]
        LightGray = 4,
        [EnumString("darkHorizontal")]
        DarkHorizontal = 5,
        [EnumString("darkVertical")]
        DarkVertical = 6,
        [EnumString("darkDown")]
        DarkDown = 7,
        [EnumString("darkUp")]
        DarkUp = 8,
        [EnumString("darkGrid")]
        DarkGrid = 9,
        [EnumString("darkTrellis")]
        DarkTrellis = 10,
        [EnumString("lightHorizontal")]
        LightHorizontal = 11,
        [EnumString("lightVertical")]
        LightVertical = 12,
        [EnumString("lightDown")]
        LightDown = 13,
        [EnumString("lightUp")]
        LightUp = 14,
        [EnumString("lightGrid")]
        LightGrid = 15,
        [EnumString("lightTrellis")]
        LightTrellis = 16,
        [EnumString("gray125")]
        Gray125 = 17,
        [EnumString("gray0625")]
        Gray0625 = 18,
    }
}
