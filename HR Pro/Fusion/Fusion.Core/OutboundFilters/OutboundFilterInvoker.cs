namespace Fusion.Core.OutboundFilters
{
    using System;
    using System.Linq;
    using StructureMap;
    using Fusion.Messages.General;
    using log4net;

    public class OutboundFilterInvoker : IOutboundFilterInvoker
    {
        private ILog Logger = LogManager.GetLogger(typeof(OutboundFilterInvoker));

        public bool Execute(FusionMessage message)
        {
            var messageType = message.GetType();
            var handlerType = typeof(IOutboundFilterHandler<>).MakeGenericType(messageType);

            var genericHandlerMatches = ObjectFactory.GetAllInstances(handlerType).Cast<IOutboundFilterHandler>();
            var handlerMatches = ObjectFactory.GetAllInstances(typeof(IOutboundFilterHandler));

            foreach (IOutboundFilterHandler handler in handlerMatches)
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