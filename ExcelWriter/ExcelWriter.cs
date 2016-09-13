using ExcelWriter.Entities;
using ExcelWriter.Parallelism;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

namespace ExcelWriter
{
    public class ExcelWriter
    {
        #region memory stuff
        
        private const long chunkSize = 20000000;

        /// <summary>
        /// Counter what increases while adding cells
        /// </summary>
        internal static long AddCellCount = 0;
        /// <summary>
        /// Counter what increases while processing cells
        /// </summary>
        internal static long RemoveCellCount = 0;

        /// <summary>
        /// While adding cells, this method is to necessary to be called. 
        /// It counts the adding of cells and if the count reached the chunksize, it waits for the count processed cells to reach the chunksize.
        /// It helps to avoid memory leaks - if chunksize set to 20000000 the maximum used megabytes of memory will be around 400.
        /// </summary>
        internal static void GCCollectIfNecessary()
        {
            if (++AddCellCount < chunkSize)
            {
                return;
            }


            if (AddCellCount == chunkSize && RemoveCellCount < chunkSize)
            {
                while (RemoveCellCount < chunkSize)
                {
                    Thread.Sleep(1);
                }
                GC.Collect();
                AddCellCount = 0;
                RemoveCellCount = 0;
            }
        }

        #endregion

        public string Name { get; set; }
        public static List<EWSheet> Sheets = new List<EWSheet>();

        EWConsumer consumer;
        private int _sheetIndex;

        internal static ConcurrentQueue<EWCell> CellQueue = new ConcurrentQueue<EWCell>();
        public static bool Finished { get; set; }

        Thread consumerThread;
        
        private const string _extension = "xlsx";
        internal static bool AddingCellsInProgress = true;

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
            ExcelWriter.AddingCellsInProgress = false;

            //waiting for the consumer thread to finish processing
            while (consumerThread.IsAlive)
            {

            }

            consumerThread.Abort();
        }
    }
}
