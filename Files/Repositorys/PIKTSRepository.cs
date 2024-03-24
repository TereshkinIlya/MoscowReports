using Data;
using Files.Abstracts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Files.Repositorys
{
    internal class PIKTSRepository : Repository<Queue<PIKTS>, PIKTS>
    {
        internal override Queue<PIKTS> Storage { get; set; }
        public PIKTSRepository()
        {
            Storage = new Queue<PIKTS>();
        }
        /// <summary>
        /// Получить элемент из хранилища
        /// </summary>
        internal override PIKTS Get()
        {
            return Storage.Dequeue();
        }
        /// <summary>
        /// Положить элемент в хранилище
        /// </summary>
        internal override void Put(PIKTS item)
        {
            Storage.Enqueue(item);
        }
    }
}
