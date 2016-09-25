namespace ExcelWriter.Entities
{
    internal struct EWCell
    {
        internal int sheetIndex;
        internal int rowIndex;
        internal int columnIndex;
        internal string value;
        internal string styleIndex;
        public EWCell(int rowIndex, int columnIndex, int sheetIndex, string styleIndex, string value)
        {
            this.rowIndex = rowIndex;
            this.columnIndex = columnIndex;
            this.value = value;
            this.sheetIndex = sheetIndex;
            this.styleIndex = styleIndex;
        }
    }
}
