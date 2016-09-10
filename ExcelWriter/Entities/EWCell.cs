namespace ExcelWriter.Entities
{
    internal struct EWCell
    {
        internal int sheetIndex;
        internal int rowIndex;
        internal int columnIndex;
        internal string value;

        public EWCell(int row, int col, string stringValue)
        {
            sheetIndex = 0;
            rowIndex = row;
            columnIndex = col;
            value = stringValue;
        }
    }
}
