using System.Collections.Generic;
using System.Threading.Tasks;

namespace TSF.SearchCandidateProviderInternal
{
    internal static class TaskEx
    {
        public static Task WhenAll(this IEnumerable<Task> tasks)
        {
            return Task.WhenAll(tasks);
        }
        public static Task<T[]> WhenAll<T>(this IEnumerable<Task<T>> tasks)
        {
            return Task.WhenAll(tasks);
        }
        public static void Forget<T>(this Task<T> task)
        {
        }
        public static void Forget(this Task task)
        {
        }
    }
}
