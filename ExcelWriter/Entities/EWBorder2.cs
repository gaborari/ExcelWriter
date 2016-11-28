using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelWriter.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWriter.Entities
{
    public class EWColor
    {
        internal System.Drawing.Color _color;

        public EWColor()
        {
            _color = System.Drawing.Color.Black;
        }

        public EWColor(System.Drawing.Color color)
        {
            this._color = color;
        }

        public string ToHexString()
        {
            return _color.R.ToString("X2") + _color.G.ToString("X2") + _color.B.ToString("X2");
        }

    }

    public class EWBorder2
    {
        internal Border oxBorder;

        public EWBorder2()
        {
            oxBorder = new Border() { Outline = true };
        }

        public void CreateLeftBorder(Borderstyles borderStyle = Borderstyles.Thin, EWColor color = null)
        {
            if (color == null)
            {
                color = new EWColor(System.Drawing.Color.Black);
            }

            LeftBorder leftBorder1 = new LeftBorder() { Style = borderStyle.As<BorderStyleValues>() };
            leftBorder1.Append(new Color() { Rgb = new HexBinaryValue() { Value = color._color.ToHexString() } });

            oxBorder.Append(leftBorder1);
        }

        public void CreateRightBorder(Borderstyles borderStyle = Borderstyles.Thin, EWColor color = null)
        {
            if (color == null)
            {
                color = new EWColor(System.Drawing.Color.Black);
            }

            RightBorder rightBorder = new RightBorder() { Style = borderStyle.As<BorderStyleValues>() };
            rightBorder.Append(new Color() { Rgb = new HexBinaryValue() { Value = color._color.ToHexString() } });

            oxBorder.Append(rightBorder);
        }

        public void CreateBottomBorder(Borderstyles borderStyle = Borderstyles.Thin, EWColor color = null)
        {
            if (color == null)
            {
                color = new EWColor(System.Drawing.Color.Black);
            }

            BottomBorder bottomBorder = new BottomBorder() { Style = borderStyle.As<BorderStyleValues>() };
            bottomBorder.Append(new Color() { Rgb = new HexBinaryValue() { Value = color._color.ToHexString() } });

            oxBorder.Append(bottomBorder);
        }

        public void CreateTopBorder(Borderstyles borderStyle = Borderstyles.Thin, EWColor color = null)
        {
            if (color == null)
            {
                color = new EWColor(System.Drawing.Color.Black);
            }

            TopBorder topBorder = new TopBorder() { Style = borderStyle.As<BorderStyleValues>() };
            topBorder.Append(new Color() { Rgb = new HexBinaryValue() { Value = color._color.ToHexString() } });

            oxBorder.Append(topBorder);
        }
    }
}
