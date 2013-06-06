

namespace Fusion.Test.SocialCare
{
    using System.IO;
    using Fusion.Core;
    using Fusion.Core.Test;
    using NServiceBus;
    using StructureMap;

    public class StartOutboundWatchers : IWantToRunAtStartup
    {
            public IBus Bus { get; set; }

            public void Run()
            {
                DirectoryUtil.EnforceDirectory("out");
                DirectoryUtil.EnforceDirectory("in");

                OutboundWatcherDefinition[] list = new OutboundWatcherDefinition[] {
                    new OutboundWatcherDefinition { PathToWatch = "StaffChangeRequest", MessageType = typeof(Fusion.Messages.SocialCare.StaffChangeRequest) },
                    new OutboundWatcherDefinition { PathToWatch = "StaffPostChangeRequest", MessageType = typeof(Fusion.Messages.SocialCare.StaffPostChangeRequest) },

                    new OutboundWatcherDefinition { PathToWatch = "ServiceUserChangeRequest", MessageType = typeof(Fusion.Messages.SocialCare.ServiceUserChangeRequest) },
                    new OutboundWatcherDefinition { PathToWatch = "ServiceUserHomeAddressChangeRequest", MessageType = typeof(Fusion.Messages.SocialCare.ServiceUserHomeAddressChangeRequest) },
                    new OutboundWatcherDefinition { PathToWatch = "ServiceUserCareDeilveryAddressChangeRequest", MessageType = typeof(Fusion.Messages.SocialCare.ServiceUserCareDeliveryAddressChangeRequest) },
                
                    new OutboundWatcherDefinition { PathToWatch = "FunderChange", MessageType = typeof(Fusion.Messages.SocialCare.FunderChangeMessage) },
                    new OutboundWatcherDefinition { PathToWatch = "FundingChange", MessageType = typeof(Fusion.Messages.SocialCare.FundingChangeMessage) },

                    new OutboundWatcherDefinition { PathToWatch = "CaresysInvoiceLineInsert", MessageType = typeof(Fusion.Messages.SocialCare.CaresysInvoiceLineInsertMessage) },

                    new OutboundWatcherDefinition { PathToWatch = "ServiceUserDailyRecordChange", MessageType = typeof(Fusion.Messages.SocialCare.ServiceUserDailyRecordChangeMessage) },
                };

                foreach(var s in list) {
                    IOutboundMessageWatcher mw = ObjectFactory.With("path").EqualTo(Path.Combine("out", s.PathToWatch))
                                    .GetInstance(typeof(OutboundMessageWatcher<>).MakeGenericType(s.MessageType)) as IOutboundMessageWatcher;

                    if (mw != null)
                    {
                        mw.Start();
                    }        
                }

            }

            public void Stop()
            {

            }
    }

   
    
}
