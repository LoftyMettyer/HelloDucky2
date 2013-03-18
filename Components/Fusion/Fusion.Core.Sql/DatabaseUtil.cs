using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Threading;
using log4net;

namespace Fusion.Core.Sql
{
    public class DatabaseUtil
    {
        static ILog Logger = LogManager.GetLogger(typeof(DatabaseUtil));

        private static void DeadlockAutoRetry(Action action)
        {
            DeadlockAutoRetry(() => { action(); return true; });
        }

        private static T DeadlockAutoRetry<T>(Func<T> func)
        {
            return DeadlockAutoRetry(func, TimeSpan.FromSeconds(1));
        }

        private static T DeadlockAutoRetry<T>(Func<T> func, TimeSpan delay)
        {
            int count = 3;

            while (true)
            {
                try
                {
                    return func();
                }
                catch (SqlException e)
                {
                    --count;
                    if (count <= 0) throw;

                    if (e.Number == 1205)
                        Logger.Debug("Deadlock, retrying", e);
                    else if (e.Number == -2)
                        Logger.Debug("Timeout, retrying", e);
                    else
                        throw;

                    Thread.Sleep(delay);
                }
            }
        }
    }
}
