using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Fusion.LogService.DatabaseAccess
{
    public class NullLogDatabase : ILogDatabase
    {
        public void AddLogEntry(Guid id, Guid? messageId, string connectorName, Guid? entityRef, DateTime time, char logLevel, string message)
        {
        }

        public void AddTranslationRecord(DateTime time, Guid? messageId, string connectorName, string translationName, Guid busRef, string localId)
        {
        }
    }
}
