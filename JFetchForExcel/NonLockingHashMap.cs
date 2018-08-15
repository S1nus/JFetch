using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using NonBlocking;
using System.Threading;
using System.Collections.Concurrent;

namespace OrionApiSdk.Utils
{

    public class NonLockingHashMap<T, U> : ConcurrentDictionary<T,U> , IEnumerable
    {

        public void Add(T key, U value)
        {
            if (ContainsKey(key))
                return;

            while (!TryAdd(key, value));
        }

        public void Update(T key, U val)
        {
            TryRemove(key,out U _temp);
            TryAdd(key, val);
        }

        public object ToList()
        {
            return Values.ToList();
        }


        public void Remove(T key)
        {
            if (key == null || !ContainsKey(key))
                return;

            while (!TryRemove(key, out U val));
        }

        // Must also implement IEnumerable.GetEnumerator, but implement as a private method.
        private IEnumerator GetEnumerator1()
        {
            return GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator1();
        }
    }

}