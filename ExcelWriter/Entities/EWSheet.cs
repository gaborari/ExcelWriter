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

        public void AddCell(int row, int col, object value)
        {
            //Only sequential writing is possible
            if (row < LastRowIndex)
            {
                return;
            }
            else
            {
                LastRowIndex = row;
                ExcelWriter.CellQueue.Enqueue(new EWCell(row, col, this.Index, value.ToString()));
                ExcelWriter.GCCollectIfNecessary();
            }
        }
    }
}
