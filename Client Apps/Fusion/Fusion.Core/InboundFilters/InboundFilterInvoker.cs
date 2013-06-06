namespace Fusion.Core.InboundFilters
{
    using System;
    using System.Linq;
    using StructureMap;
    using Fusion.Messages.General;
using log4net;

    public class InboundFilterInvoker : IInboundFilterInvoker
    {
        private ILog Logger = LogManager.GetLogger(typeof(InboundFilterInvoker));
        
        public bool Execute(FusionMessage message)
        {
            var messageType = message.GetType();
            var handlerType = typeof(IInboundFilterHandler<>).MakeGenericType(messageType);

            var genericHandlerMatches = ObjectFactory.GetAllInstances(handlerType).Cast<IInboundFilterHandler>();
            var handlerMatches = ObjectFactory.GetAllInstances(typeof(IInboundFilterHandler));

            foreach (IInboundFilterHandler handler in handlerMatches)
            {
                bool shouldContinue = handler.Handle(message);

                Logger.InfoFormat("Invoked filter {0}, result = {1}", handler.GetType().Name, shouldContinue);

                if (!shouldContinue)
                {
                    return false;
                }
            }

            foreach (var handler in genericHandlerMatches)
            {
                bool shouldContinue = handler.Handle(message);

                Logger.InfoFormat("Invoked filter {0}, result = {1}", handler.GetType().Name, shouldContinue);

                if (!shouldContinue)
                {
                    return false;
                }
            }

            return true;
        }
    }
}