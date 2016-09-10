using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelWriter.Tests
{
    [TestClass]
    public class Tests
    {
        [TestMethod]
        public void MainTest()
        {
            var excelWriter = new ExcelWriter();
            var elsoSheet = excelWriter.AddSheet("firstsheet");

            for (int i = 0; i < 500000; i++)
            {
                elsoSheet.AddCell(1, 1, "testvalue"+ i);
            }


            excelWriter.Generate();
        }

    }
}
