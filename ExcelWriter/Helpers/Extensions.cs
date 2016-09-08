using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelWriter.Entities;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
namespace ExcelWriter.Helpers
{
    internal static class Extensions
    {

        /// <summary>
        /// Get the column name from an int number
        /// </summary>
        /// <param name="columnNumber">the index of column</param>
        /// <returns>the name of column</returns>
        public static string GetColumnName(this int columnNumber)
        {
            int temp = columnNumber;
            string columnName = string.Empty;
            int modulo;

            while (temp > 0)
            {
                modulo = (temp - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                temp = (int)((temp - modulo) / 26);
            }

            return columnName;
        }

        /// <summary>
        /// Gets the row index from a cell name
        /// </summary>
        /// <param name="cellName">the name of the cell</param>
        /// <returns>the row index</returns>
        internal static uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }

        /// <summary>
        /// Writes all merged cells after to the end of the xml document
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="mergedCells"></param>
        internal static void WriteMergedCells(this OpenXmlWriter writer, IEnumerable<EWMergedCell> mergedCells)
        {

            if (mergedCells.Any())
            {
                MergeCells mergeCells = new MergeCells();

                writer.WriteStartElement(mergeCells);

                foreach (var mg in mergedCells)
                {
                    writer.WriteElement(mg.MergedCell);
                }

                writer.WriteEndElement();
            }
        }
    }
}
