using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenHRTestToLive.Extensions
{
  public static class CollectionExtension
  {
    public static void AddUniqueBy<T>(this ICollection<T> source, Func<T, T, bool> predicate, IEnumerable<T> items)
    {
      foreach (var item in items)
      {
        var existsInSource = source.Any(s => predicate(s, item));
        if (!existsInSource) source.Add(item);
      }
    }

  }
}
