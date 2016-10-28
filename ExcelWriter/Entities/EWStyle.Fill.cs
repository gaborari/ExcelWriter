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

        public EWFill(EWPattern pattern, System.Drawing.Color color, string foreGroundColor = null)
        {
            if (!string.IsNullOrEmpty(foreGroundColor))
            {
                _oxFill = new Fill(                                                           // Index 2 – The yellow fill.
                            new PatternFill(
                                new ForegroundColor() { Rgb = new HexBinaryValue() { Value = foreGroundColor } }
                            ) { PatternType = pattern.As<PatternValues>() }
                        );
                return;
            }

            _oxFill = new Fill();
            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = "FF0070C0" };
            BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill4.Append(foregroundColor2);
            patternFill4.Append(backgroundColor2);
            _oxFill.Append(patternFill4);

            //_oxFill = new Fill(                                                           // Index 2 – The yellow fill.
            //         new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } })
            //         {
            //             PatternType = PatternValues.Solid,
            //             BackgroundColor = new BackgroundColor { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } }
            //         });
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
