using System.Collections.Generic;
using System.Linq;

namespace OutlookSync
{
    public static class EnumerableExtensions
    {
        public static IEnumerable<IEnumerable<T>> Chunk<T>(this IEnumerable<T> source, int chunk_size)
        {
            while (source.Any())
            {
                yield return source.Take(chunk_size);
                source = source.Skip(chunk_size);
            }
        }
    }
}
