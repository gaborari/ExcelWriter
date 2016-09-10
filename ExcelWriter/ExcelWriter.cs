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
        private List<EWSheet> _sheets = new List<EWSheet>();

        Object _lockObj = new object();
        private ConcurrentQueue<EWCell> _cellQueue = new ConcurrentQueue<EWCell>();
        EWProducer producer;
        EWConsumer consumer;
        EWThreadKiller threadKiller = new EWThreadKiller();
        EWProcessingFinished finished = new EWProcessingFinished();
        private int _sheetIndex;

        Thread consumerThread;
        Thread producerThread;
        public ExcelWriter()
        {

            producer = new EWProducer(_cellQueue, _lockObj, threadKiller, finished);
            consumer = new EWConsumer(_cellQueue, _lockObj, "deprecated", threadKiller, finished);

            consumerThread = new Thread(consumer.consume);
            consumerThread.Start();
            Debug.WriteLine("Consumer started");


            Debug.WriteLine("Producer start");
            producerThread = new Thread(() => producer.produce(new EWCell()));
            producerThread.Start();
            Debug.WriteLine("Producer started");

           
        }

        public EWSheet AddSheet(string name)
        {
            var sheet = new EWSheet() { Name = name, Index = _sheetIndex++ };
            this._sheets.Add(sheet);
            
            sheet.CellAdded += CellAdded;

            return sheet;
        }

        private void CellAdded(object sender, CellAddedEventArgs e)
        {
            producer.produce(e.Cell);
        }

        public void Generate()
        {
            producer.Stop();
            
            
            //waiting for the consumer thread to finish processing
            while (consumerThread.IsAlive)
            {

            }

            consumerThread.Abort();
            producerThread.Abort();
        }
    }
}
