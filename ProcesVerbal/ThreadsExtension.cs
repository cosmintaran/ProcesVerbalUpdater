using System;
using System.Threading.Tasks;

namespace ProcesVerbal
{
    internal static class ThreadsExtension
    {
        public static async void Await(this Task task, Action onExecuted = null, Action<Exception> onException = null)
        {
            try
            {
                await task;
                onExecuted?.Invoke();
            }
            catch(Exception ex)
            {
                onException?.Invoke(ex);
            }
        }
    }
}
