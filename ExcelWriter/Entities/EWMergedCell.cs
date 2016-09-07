using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelWriter.Helpers;
namespace ExcelWriter.Entities
{
    internal struct EWMergedCell
    {
        internal int FromRow;
        internal int ToRow;
        internal int FromCol;
        internal int ToCol;
        internal MergeCell MergedCell;

        public EWMergedCell(int fromRow, int toRow, int fromCol, int toCol)
        {
            FromRow = fromRow;
            ToRow = toRow;
            FromCol = fromCol;
            ToCol = toCol;

            MergedCell = new MergeCell()
            {
                Reference = new StringValue(string.Format("{0}{1}:{2}{3}", fromCol.GetColumnName(), fromRow, toCol.GetColumnName(), toRow))
            };
        }
    }
}
