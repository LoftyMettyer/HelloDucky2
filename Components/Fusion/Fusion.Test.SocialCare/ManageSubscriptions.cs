using System;
using System.Configuration;
using Fusion.Core.Test.Configuration;
using log4net;
using NServiceBus;

namespace Fusion.Test.SocialCare
{
    /// <summary>
    /// Showing how to manage subscriptions from a configuration file
    /// </summary>
    class ManageSubscriptions : IWantToRunAtStartup
    {
        public IBus Bus { get; set; }

        private static readonly ILog Logger = LogManager.GetLogger(typeof(ManageSubscriptions));

        public void Run()
        {
            SubscriptionsSection section = ConfigurationManager.GetSection("SubscriptionConfig") as SubscriptionsSection;

            if (section != null)
            {
                foreach (SubscriptionTypeElement article in section.Subscriptions)
                {
                    Logger.Info("Attempting to subscribe to " + article.Type);

                    Type subscribeType = Type.GetType(article.Type);
                    if (subscribeType == null)
                    {
                        Logger.Error("Cannont find type to subscribe to " + article.Type);

                    }
                    else
                    {
                        Bus.Subscribe(subscribeType);
                    }
                }
            }

        }

        public void Stop()
        {

        }
    }
}
