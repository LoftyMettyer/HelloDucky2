
namespace Fusion.Core.Sql
{
    using System;
    using Fusion.Core.Logging;

    public class LoggingBusRefTranslatorDecorator : IBusRefTranslator
    {
        public LoggingBusRefTranslatorDecorator(string community, IBusRefTranslator busTranslator, IFusionLogService fusionLogger)
        {
            this.busTranslator = busTranslator;
            this.fusionLogger = fusionLogger;
            this.community = community;
        }

        string community;
        IBusRefTranslator busTranslator;
        IFusionLogService fusionLogger;

        public string GetLocalRef(string translationName, Guid busRef)
        {
            return this.busTranslator.GetLocalRef(translationName, busRef);
        }

        public void SetBusRef(string translationName, string localId, Guid busRef)
        {
            this.fusionLogger.LogRefTranslationTransactional(this.community, translationName, busRef, localId);

            this.busTranslator.SetBusRef(translationName, localId, busRef);
        }

        public BusTranslationResults TryGetBusRef(string translationName, string localId, bool canCreate = false)
        {
            BusTranslationResults results = busTranslator.TryGetBusRef(translationName, localId, canCreate);

            if (results.BusRefNewlyCreated)
            {
                this.fusionLogger.LogRefTranslationTransactional(this.community, translationName, results.Value, localId);
            }

            return results;
        }
    }
}
