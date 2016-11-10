namespace ExcelWriter.Entities
{
    internal struct EWCell
    {
        internal int sheetIndex;
        internal int rowIndex;
        internal int columnIndex;
        internal string value;
        internal string styleIndex;
        internal string cellType;
        public EWCell(int rowIndex, int columnIndex, int sheetIndex, string styleIndex, string value, string cellType)
        {
            this.rowIndex = rowIndex;
            this.columnIndex = columnIndex;
            this.value = value;
            this.sheetIndex = sheetIndex;
            this.styleIndex = styleIndex;
            this.cellType = cellType;
        }
    }

    public enum CellType
    {
        General = 0,
        Numeric = 1,
        Date = 2,
        DateTime = 3,
        String = 4
    }
}
