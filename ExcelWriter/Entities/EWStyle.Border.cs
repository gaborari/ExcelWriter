﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelWriter.Entities;
using ExcelWriter.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelWriter
{
    [Flags]
    public enum BorderType
    {
        Left = 1,
        Right = 2,
        Top = 4,
        Bottom = 8,
        All = Left | Right | Top | Bottom
    }

    public class EWBorder
    {
        internal Border oxBorder
        {
            get
            {
                if (_borders.Any())
                {
                    _borders = _borders.DistinctBy(x => x.GetType().Name).ToList();

                    //Must order the borders - left, right, top, bottom
                    //TODO: refactor
                    _borders = _borders.OrderBy(x => x.GetType().Name == BorderType.Left.ToString() + "Border" ? 0
                    : x.GetType().Name == BorderType.Right.ToString() + "Border" ? 1
                    : x.GetType().Name == BorderType.Top.ToString() + "Border" ? 2
                    : 3).ToList();

                    _oxBorder.Append(_borders);

                }
                return _oxBorder;
            }
        }

        private Border _oxBorder;

        private List<BorderPropertiesType> _borders = new List<BorderPropertiesType>();

        public EWBorder()
        {
            _oxBorder = new Border() { Outline = true };
        }

        private void AddBorder<T>(Borderstyles borderStyle, EWColor color)
        {
            var border = Activator.CreateInstance<T>() as BorderPropertiesType;
            border.Style = borderStyle.As<BorderStyleValues>();
            border.Append(new Color() { Rgb = new HexBinaryValue() { Value = color.HexString } });
            _borders.Add(border);
        }

        public void CreateBorder(BorderType type, Borderstyles borderStyle = Borderstyles.Thin, EWColor color = null)
        {
            if (color == null)
            {
                color = new EWColor(System.Drawing.Color.Black);
            }

            if (type.HasFlag(BorderType.Left))
            {
                AddBorder<LeftBorder>(borderStyle, color);
            }

            if (type.HasFlag(BorderType.Right))
            {
                AddBorder<RightBorder>(borderStyle, color);
            }

            if (type.HasFlag(BorderType.Top))
            {
                AddBorder<TopBorder>(borderStyle, color);
            }

            if (type.HasFlag(BorderType.Bottom))
            {
                AddBorder<BottomBorder>(borderStyle, color);
            }
        }
    }

    public enum Borderstyles
    {
        None = 0,
        Thin = 1,
        Medium = 2,
        Dashed = 3,
        Dotted = 4,
        Thick = 5,
        Double = 6,
        Hair = 7,
        MediumDashed = 8,
        DashDot = 9,
        MediumDashDot = 10,
        DashDotDot = 11,
        MediumDashDotDot = 12,
        SlantDashDot = 13
    }
}
