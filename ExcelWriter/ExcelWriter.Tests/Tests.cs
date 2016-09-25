using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System.Diagnostics;

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
            

            int rowIndex = 1;
            int colIndex = 1;
            for (int i = 0; i < 5; i++)
            {
                var sheet = excelWriter.AddSheet("sheet" + i);

                foreach (var row in rows)
                {
                    foreach (var col in row.Columns)
                    {
                        sheet.AddCell(rowIndex, colIndex++, "default", "testvalue");
                    }
                    rowIndex++;
                }

            }
            Stopwatch sp = new Stopwatch(); sp.Start();

            excelWriter.Generate();

            sp.Stop();
        }

    }

    public class TestData
    {
        public static IEnumerable<TestRow> GetData() 
        {
            for (int i = 0; i < 100; i++)
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
