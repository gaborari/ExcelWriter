using ExcelWriter.Entities;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelWriter.Parallelism
{
    internal class EWConsumer
    {
        string name;
        private ConcurrentQueue<EWCell> _cellQueue;
        private object _lockObj;
        private string p;
        private EWThreadKiller threadKiller;
        private EWProcessingFinished finished;

        public EWConsumer(ConcurrentQueue<EWCell> _cellQueue, object _lockObj, string p, EWThreadKiller threadKiller, EWProcessingFinished finished)
        {
            // TODO: Complete member initialization
            this._cellQueue = _cellQueue;
            this._lockObj = _lockObj;
            this.p = p;
            this.threadKiller = threadKiller;
            this.finished = finished;
        }

        public void consume()
        {
            EWCell item;
            while (true)
            {
                lock (_lockObj)
                {
                    if (!_cellQueue.IsEmpty)
                    {
                        _cellQueue.TryDequeue(out item);
                        //Debug.WriteLine(item.value);
                        //break;
                    }
                    else
                    {
                        if (this.threadKiller.Killed == true)
                        {
                            break;
                        }

                        Monitor.Wait(_lockObj);
                        continue;
                    }
                }
            }

            lock (_lockObj)
            {
                finished.EWFinished = true;
            }
        }
    }
}
