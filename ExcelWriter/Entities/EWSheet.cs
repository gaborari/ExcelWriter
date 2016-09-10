using ExcelWriter.Parallelism;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWriter.Entities
{
    public class EWSheet
    {
        public string Name;
        internal int Index { get; set; }
        internal int LastRowIndex { get; set; }

        internal virtual void OnCellAdded(CellAddedEventArgs e)
        {
            EventHandler<CellAddedEventArgs> handler = CellAdded;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        internal event EventHandler<CellAddedEventArgs> CellAdded;

        public void AddCell(int row, int col, object value)
        {
            if (row < LastRowIndex)
            {
                return;
            }
            else
            {
                LastRowIndex = row;
                var cell = new EWCell(row, col, value.ToString());
                cell.sheetIndex = this.Index;
                CellAddedEventArgs eventarg = new CellAddedEventArgs();
                eventarg.SheetIndex = this.Index;
                eventarg.Cell = cell;
                OnCellAdded(eventarg);
            }
        }
    }

    internal class CellAddedEventArgs : EventArgs
    {
        internal int SheetIndex { get; set; }
        internal EWCell Cell { get; set; }
    }
}
