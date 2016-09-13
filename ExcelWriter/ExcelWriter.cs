using ExcelWriter.Entities;
using ExcelWriter.Parallelism;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelWriter
{
    public class ExcelWriter
    {
        public string Name { get; set; }
        public static List<EWSheet> Sheets = new List<EWSheet>();

        Object _lockObj = new object();
        
        EWConsumer consumer;
        EWThreadKiller threadKiller = new EWThreadKiller();
        EWProcessingFinished finished = new EWProcessingFinished();
        private int _sheetIndex;

        internal static ConcurrentQueue<EWCell> CellQueue = new ConcurrentQueue<EWCell>();
        public static bool Finished { get; set; }

        Thread consumerThread;
        
        private const string _extension = "xlsx";

        public ExcelWriter(string fileName = null)
        {
            if(string.IsNullOrEmpty(fileName))
            {
                Name = string.Format("{0}.{1}", Guid.NewGuid().ToString(), _extension);
            }

            consumer = new EWConsumer(Name);

            consumerThread = new Thread(consumer.consume);
            consumerThread.Start();
            Debug.WriteLine("Consumer started");
        }

        public EWSheet AddSheet(string name)
        {
            var sheet = new EWSheet() { Name = name, Index = _sheetIndex++ };
            Sheets.Add(sheet);
            return sheet;
        }

        public void Generate()
        {
            //waiting for the consumer thread to finish processing
            while (consumerThread.IsAlive)
            {

            }

            consumerThread.Abort();
        }
    }
}
