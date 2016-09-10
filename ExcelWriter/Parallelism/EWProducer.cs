using ExcelWriter.Entities;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelWriter.Parallelism
{
    internal class EWProducer
    {
        int seq = 0;
        private ConcurrentQueue<EWCell> _cellQueue;
        private object _lockObj;
        private EWThreadKiller threadKiller;
        private EWProcessingFinished finished;

        private bool _finished;
        public EWProducer(ConcurrentQueue<EWCell> _cellQueue, object _lockObj, EWThreadKiller threadKiller, EWProcessingFinished finished)
        {
            this._cellQueue = _cellQueue;
            this._lockObj = _lockObj;
            this.threadKiller = threadKiller;
            this.finished = finished;
        }

        public void produce(EWCell cell)
        {
            lock (_lockObj)
            {

                _cellQueue.Enqueue(cell);
                if (_cellQueue.Count == 1)
                { // first
                    Monitor.PulseAll(_lockObj);
                }
            }

            lock (_lockObj)
            {
                Monitor.PulseAll(_lockObj);
            }


            if(_finished == true)
            {
                lock (_lockObj)
                {
                    threadKiller.Killed = true;
                    Monitor.PulseAll(_lockObj);
                }

                while (finished.EWFinished != true)
                {

                }
                Console.WriteLine("CONUSMER finished");
            }
        }

        internal void Stop()
        {
            lock (_lockObj)
            {
                _finished = true;
                threadKiller.Killed = true;
                Monitor.PulseAll(_lockObj);
            }
        }
    }
}
