namespace ExcelWriter.Entities
{
    public class EWSheet
    {
        public string Name;
        internal int Index { get; set; }
        private int _lastRowIndex { get; set; }
        private int _lastColIndex { get; set; }
        public void AddCell(int row, int col, string styleSelector, object value, CellType cellType = CellType.General)
        {
            //Only sequential writing is possible
            if (row < _lastRowIndex)
            {
                return;
            }

            if (col < _lastColIndex)
            {
                return;
            }

            string styleIndex = GetStyleIndex(styleSelector);

            _lastRowIndex = row;
            _lastColIndex = col;
            ExcelWriter.CellQueue.Enqueue(new EWCell(row, col, this.Index, styleIndex, value.ToString(), "str"));

            if (!ExcelWriter.DisableMemoryRestriction)
            {
                ExcelWriter.GCCollectIfNecessary();
            }
        }

        //TODO Handle properly
        private static string GetType(CellType cellType)
        {
            switch (cellType)
            {
                
                case CellType.Numeric:
                    return "str";
                case CellType.Date:
                    return "str";
                case CellType.DateTime:
                    return "str";
                case CellType.General:
                case CellType.String:
                default:
                    return "str";
            }
        }

        private string GetStyleIndex(string styleSelector)
        {
            if (EWStyle.selectors.Count == 0)
            {
                //return the default style index;
                return "0";
            }

            return EWStyle.selectors[styleSelector];
        }

        private void AddEmptyRows(int currentRow)
        {
            if (_lastRowIndex == currentRow)
            {
                return;
            }

            var gap = currentRow - _lastRowIndex;
            if (gap <= 1)
            {
                return;
            }

            gap--;
            for (int i = 0; i < gap; i++)
            {

            }
        }

        private void AddEmptyCols(int currentCol)
        {
            if (_lastColIndex == currentCol)
            {
                return;
            }

            var gap = currentCol - _lastColIndex;
            if (gap <= 1)
            {
                return;
            }

            gap--;
            for (int i = 0; i < gap; i++)
            {

            }
        }
    }
}
