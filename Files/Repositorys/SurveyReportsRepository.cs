using Files.Abstracts;
using Data;

namespace Files.Repositorys
{
    internal class SurveyReportsRepository : Repository<Queue<UnderwaterCrossing>, UnderwaterCrossing>
    {
        internal override Queue<UnderwaterCrossing> Storage { get; set; }
        public SurveyReportsRepository()
        {
            Storage = new Queue<UnderwaterCrossing>();
        }
        /// <summary>
        /// Получить элемент из хранилища
        /// </summary>
        internal override UnderwaterCrossing Get()
        {
            return Storage.Dequeue();
        }
        /// <summary>
        /// Положить элемент в хранилище
        /// </summary>
        internal override void Put(UnderwaterCrossing item)
        {
            Storage.Enqueue(item);
        }
    }
}
