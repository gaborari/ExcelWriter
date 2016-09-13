namespace ExcelWriter.Entities
{
    internal struct EWCell
    {
        internal int sheetIndex;
        internal int rowIndex;
        internal int columnIndex;
        internal string value;

        public EWCell(int rowIndex, int columnIndex, int sheetIndex, string value)
        {
            this.rowIndex = rowIndex;
            this.columnIndex = columnIndex;
            this.value = value;
            this.sheetIndex = sheetIndex;
        }
    }
}
