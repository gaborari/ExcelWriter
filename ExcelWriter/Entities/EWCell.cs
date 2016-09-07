namespace ExcelWriter.Entities
{
    internal struct EWCell
    {
        internal int rowIndex;
        internal int columnIndex;
        internal string value;

        public EWCell(int row, int col, string stringValue)
        {
            rowIndex = row;
            columnIndex = col;
            value = stringValue;
        }
    }
}
