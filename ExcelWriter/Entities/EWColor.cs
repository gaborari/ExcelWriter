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

        public string HexString
        {
            get
            {
                return this.ToHexString();
            }
        }


        public string ToHexString()
        {
            return _color.R.ToString("X2") + _color.G.ToString("X2") + _color.B.ToString("X2");
        }

    }

}
