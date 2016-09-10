using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace ExcelWriter.Tests
{
    [TestClass]
    public class Tests
    {
        public List<TestRow> rows = TestData.GetData().ToList();

        [TestMethod]
        public void MainTest()
        {
            var excelWriter = new ExcelWriter();
            var elsoSheet = excelWriter.AddSheet("firstsheet");

            int rowIndex = 1;
            int colIndex = 1;
            foreach (var row in rows)
            {
                foreach (var col in row.Columns)
                {
                    elsoSheet.AddCell(rowIndex, colIndex++, "testvalue");
                }
                rowIndex++;
            }

            excelWriter.Generate();
        }

    }

    public class TestData
    {
        public static IEnumerable<TestRow> GetData() 
        {
            for (int i = 0; i < 100000; i++)
            {
                yield return new TestRow { Columns = TestRow.GetColumns() } ;   
            }
        }
    }

    public class TestRow
    {
        public IEnumerable<TestColumn> Columns { get; set; }


        public static IEnumerable<TestColumn> GetColumns()
        {
            for (int i = 0; i < 100; i++)
            {
                yield return new TestColumn { Value = i.ToString() };
            }
        }
    }

    public class TestColumn
    {
        public string Value { get; set; }
    }
}
