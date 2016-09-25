using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelWriter.Exceptions;
using ExcelWriter.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelWriter
{
    public class EWBorder 
    {
        private bool _isDefaultBorder = true;
        private bool _isRegisteredBorder = false;

        private Border _oxBorder;

        #region Border visibilities
        #region Left border visibility
        private bool _leftBorderVisible = true;
        public bool LeftBorderVisible
        {
            get
            {
                return _leftBorderVisible;
            }
            set
            {
                if (_isRegisteredBorder)
                {
                    throw new Exception("Cannot modify border after initialization");
                }

                if (_isDefaultBorder)
                {
                    _isDefaultBorder = false;
                }

                _leftBorderVisible = value;
            }
        }
        #endregion

        #region Right border visibility
        private bool _rightBorderVisible = true;
        public bool RightBorderVisible
        {
            get
            {
                return _rightBorderVisible;
            }
            set
            {
                if (_isRegisteredBorder)
                {
                    throw new Exception("Cannot modify border after initialization");
                }

                if (_isDefaultBorder)
                {
                    _isDefaultBorder = false;
                }

                _rightBorderVisible = value;
            }
        }
        #endregion

        #region Top border visibility
        private bool _topBorderVisible = true;
        public bool TopBorderVisible
        {
            get
            {
                return _topBorderVisible;
            }
            set
            {
                if (_isRegisteredBorder)
                {
                    throw new Exception("Cannot modify border after initialization");
                }

                if (_isDefaultBorder)
                {
                    _isDefaultBorder = false;
                }

                _topBorderVisible = value;
            }
        }
        #endregion

        #region Bottom border visibility
        private bool _bottomBorderVisible = true;
        public bool BottomBorderVisible
        {
            get
            {
                return _bottomBorderVisible;
            }
            set
            {
                if (_isRegisteredBorder)
                {
                    throw new Exception("Cannot modify border after initialization");
                }

                if (_isDefaultBorder)
                {
                    _isDefaultBorder = false;
                }

                _bottomBorderVisible = value;
            }
        }
        #endregion
        #endregion

        #region Border colours

        #region Left border colour

        private string _leftBorderColour;
        public string LeftBorderColour
        {
            set
            {
                if (_isRegisteredBorder)
                {
                    throw new Exception("Cannot modify border after initialization");
                }

                if (_isDefaultBorder)
                {
                    _isDefaultBorder = false;
                }

                _leftBorderColour = value;
            }
        }

        #endregion

        #region Right border colour

        private string _rightBorderColour;
        public string RightBorderColour
        {
            set
            {
                if (_isRegisteredBorder)
                {
                    throw new Exception("Cannot modify border after initialization");
                }

                if (_isDefaultBorder)
                {
                    _isDefaultBorder = false;
                }

                _rightBorderColour = value;
            }
        }

        #endregion

        #region Top border colour

        private string _topBorderColour;
        public string TopBorderColour
        {
            set
            {
                if (_isRegisteredBorder)
                {
                    throw new Exception("Cannot modify border after initialization");
                }

                if (_isDefaultBorder)
                {
                    _isDefaultBorder = false;
                }

                _topBorderColour = value;
            }
        }

        #endregion

        #region Bottom border colour

        private string _bottomBorderColour;
        public string BottomBorderColour
        {
            set
            {
                if (_isRegisteredBorder)
                {
                    throw new Exception("Cannot modify border after initialization");
                }

                if (_isDefaultBorder)
                {
                    _isDefaultBorder = false;
                }

                _bottomBorderColour = value;
            }
        }

        #endregion

        #endregion

        #region Border styles

        #region Left border style
        private Borderstyles? _leftBorderStyle;
        public Borderstyles LeftBorderStyle
        {
            set
            {
                if (_isRegisteredBorder)
                {
                    throw new Exception("Cannot modify border after initialization");
                }

                if (_isDefaultBorder)
                {
                    _isDefaultBorder = false;
                }

                _leftBorderStyle = value;
            }
        }
        #endregion

        #region Right border style
        private Borderstyles? _rightBorderStyle;
        public Borderstyles RightBorderStyle
        {
            set
            {
                if (_isRegisteredBorder)
                {
                    throw new Exception("Cannot modify border after initialization");
                }

                if (_isDefaultBorder)
                {
                    _isDefaultBorder = false;
                }

                _rightBorderStyle = value;
            }
        }
        #endregion

        #region Top border style
        private Borderstyles? _topBorderStyle;
        public Borderstyles TopBorderStyle
        {
            set
            {
                if (_isRegisteredBorder)
                {
                    throw new Exception("Cannot modify border after initialization");
                }

                if (_isDefaultBorder)
                {
                    _isDefaultBorder = false;
                }

                _topBorderStyle = value;
            }
        }
        #endregion

        #region Bottom border style
        private Borderstyles? _bottomBorderStyle;
        public Borderstyles BottomBorderStyle
        {
            set
            {
                if (_isRegisteredBorder)
                {
                    throw new Exception("Cannot modify border after initialization");
                }

                if (_isDefaultBorder)
                {
                    _isDefaultBorder = false;
                }

                _bottomBorderStyle = value;
            }
        }
        #endregion

        #endregion

        internal Border GetOpenXmlBorder()
        {
            _isRegisteredBorder = true;

            if (this._isDefaultBorder)
            {
                return new Border(                                                         // Index 0 – The default border.
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder());
            }

            LeftBorder leftborder = null;
            RightBorder rightborder = null;
            TopBorder topborder = null;
            BottomBorder bottomborder = null;

            #region initializing the sides if necessary
            //if (_leftBorderVisible)
            //{
            //    leftborder = new LeftBorder();

            //    #region Setting colours to borders
            //    if (_leftBorderColour != null)
            //    {
            //        leftborder.Color = new Color() { Rgb = new HexBinaryValue() { Value = _leftBorderColour } };
            //    }
            //    #endregion

            //    #region Setting styles to borders
            //    if (_leftBorderStyle != null)
            //    {
            //        leftborder.Style = _leftBorderStyle.As<BorderStyleValues>();
            //    }
            //    #endregion
            //}

            //if (_rightBorderVisible)
            //{
            //    rightborder = new RightBorder();

            //    #region Setting colours to borders
            //    if (_rightBorderColour != null)
            //    {
            //        rightborder.Color = new Color() { Rgb = new HexBinaryValue() { Value = _rightBorderColour } };
            //    }
            //    #endregion

            //    #region Setting styles to borders
            //    if (_rightBorderStyle != null)
            //    {
            //        rightborder.Style = _rightBorderStyle.As<BorderStyleValues>();
            //    }
            //    #endregion
            //}


            //if (_topBorderVisible)
            //{
            //    topborder = new TopBorder();

            //    #region Setting colours to borders
            //    if (_topBorderColour != null)
            //    {
            //        topborder.Color = new Color() { Rgb = new HexBinaryValue() { Value = _topBorderColour } };
            //    }
            //    #endregion

            //    #region Setting styles to borders
            //    if (_topBorderStyle != null)
            //    {
            //        topborder.Style = _topBorderStyle.As<BorderStyleValues>();
            //    }
            //    #endregion
            //}


            //if (_bottomBorderVisible)
            //{
            //    bottomborder = new BottomBorder();

            //    #region Setting colours to borders
            //    if (_bottomBorderColour != null)
            //    {
            //        bottomborder.Color = new Color() { Rgb = new HexBinaryValue() { Value = _bottomBorderColour } };
            //    }
            //    #endregion

            //    #region Setting styles to borders
            //    if (_bottomBorderStyle != null)
            //    {
            //        bottomborder.Style = _bottomBorderStyle.As<BorderStyleValues>();
            //    }
            //    #endregion
            //}
            #endregion

            InitializeBorder(_leftBorderVisible, _leftBorderColour, _leftBorderStyle, leftborder, typeof(LeftBorder));
            InitializeBorder(_rightBorderVisible, _rightBorderColour, _rightBorderStyle, rightborder, typeof(LeftBorder));
            InitializeBorder(_topBorderVisible, _topBorderColour, _topBorderStyle, topborder, typeof(LeftBorder));
            InitializeBorder(_bottomBorderVisible, _bottomBorderColour, _bottomBorderStyle, bottomborder, typeof(LeftBorder));

            var a = leftborder == null;
            var b = rightborder == null;
            var c = topborder == null;
            var d = bottomborder == null;

            #region odl
            //if (a && b && c && d)     return new Border(leftborder, rightborder, topborder, bottomborder);
            //if (a && b && c && !d)    return new Border(leftborder, rightborder, topborder);

            //if (a && b && !c && d)    return new Border(leftborder, rightborder, bottomborder);
            //if (a && b && !c && !d)   return new Border(leftborder, rightborder);

            //if (a && !b && c && d)    return new Border(leftborder, topborder, bottomborder);
            //if (a && !b && !c && d)   return new Border(leftborder, bottomborder);
            //if (a && !b && c && !d)   return new Border(leftborder, topborder);
            //if (a && !b && !c && !d)  return new Border(leftborder);

            //if (!a && b && c && d)    return new Border(rightborder, topborder, bottomborder);
            //if (!a && b && c && !d)   return new Border(rightborder, topborder);
            //if (!a && b && !c && d)   return new Border(rightborder, bottomborder);
            //if (!a && !b && c && d)   return new Border(topborder, bottomborder);
            //if (!a && b && !c && !d)  return new Border(rightborder, bottomborder);
            //if (!a && !b && c && !d)  return new Border(topborder);
            //if (!a && !b && !c && d)  return new Border(bottomborder);
            //if (!a && !b && !c && !d) return new Border(); 
            #endregion

            var tempArray = (new List<BorderPropertiesType> {leftborder, rightborder, topborder, bottomborder }).Where(x => x != null).ToArray();

            return new Border(tempArray);
        }

        private void InitializeBorder(bool isVisible, string borderColour, Borderstyles? style, BorderPropertiesType border, Type borderType)
        {
            if (isVisible)
            {
                border = (BorderPropertiesType)Activator.CreateInstance(borderType);

                #region Setting colours to borders
                if (borderColour != null)
                {
                    border.Color = new Color() { Rgb = new HexBinaryValue() { Value = borderColour } };
                }
                #endregion

                #region Setting styles to borders
                if (style != null)
                {
                    border.Style = style.As<BorderStyleValues>();
                }
                #endregion
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
