using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Files.Abstracts
{
    internal abstract class Repository<T,K>where T : class 
                                           where K : class
    {
        internal abstract T Storage { get; set; }
        internal abstract void Put(K item);
        internal abstract K Get();
    }
}
