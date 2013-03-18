using log4net.Appender;
using log4net.Core;
using NServiceBus;

namespace Fusion.Core.LoggingHandlers
{
    /// <summary>
    /// Handles logging configuration for the integration profile.
    /// </summary>
    public class IntegrationLoggingHandler : IConfigureLoggingForProfile<Fusion.Core.Integration>
    {
        void IConfigureLogging.Configure(IConfigureThisEndpoint specifier)
        {
            SetLoggingLibrary.Log4Net<ColoredConsoleAppender>(null,
                a =>
                {
                    NServiceBus.Hosting.Windows.LoggingHandlers.LiteLoggingHandler.PrepareColors(a);
                    a.Threshold = Level.Info;
                }
            );
        }
    }
}