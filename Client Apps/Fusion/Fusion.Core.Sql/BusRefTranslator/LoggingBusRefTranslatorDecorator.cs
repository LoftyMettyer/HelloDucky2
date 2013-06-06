using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using StructureMap.Attributes;
using Fusion.Core.Logging;

namespace Fusion.Core.Sql
{
    public class LoggingBusRefTranslatorDecorator : IBusRefTranslator
    {


        public LoggingBusRefTranslatorDecorator(IBusRefTranslator busTranslator, IFusionLogService fusionLogger)
        {
            this.busTranslator = busTranslator;
            this.fusionLogger = fusionLogger;
        }

        IBusRefTranslator busTranslator;
        IFusionLogService fusionLogger;

        public string GetLocalRef(string translationName, Guid busRef)
        {
            return busTranslator.GetLocalRef(translationName, busRef);
        }

        public void SetBusRef(string translationName, string localId, Guid busRef)
        {
            fusionLogger.LogRefTranslationTransactional(translationName, busRef, localId);

            busTranslator.SetBusRef(translationName, localId, busRef);
        }

        public BusTranslationResults TryGetBusRef(string translationName, string localId, bool canCreate = false)
        {
            BusTranslationResults results = busTranslator.TryGetBusRef(translationName, localId, canCreate);

            if (results.BusRefNewlyCreated)
            {
                fusionLogger.LogRefTranslationTransactional(translationName, results.Value, localId);
            }

            return results;
        }
    }
}
