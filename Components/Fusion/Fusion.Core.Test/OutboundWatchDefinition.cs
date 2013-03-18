
namespace Fusion.Core.Test
{
    using System;
    using Fusion.Core.MessageSenders;

    public class OutboundWatcherDefinition
    {
        public string PathToWatch
        {
            get;
            set;
        }

        public Type MessageType
        {
            get;
            set;
        }

        public IFusionXmlMetadataExtract MetadataExtrator
        {
            get;
            set;
        }

        public IMessageSender MessageSender
        {
            get;
            set;
        }
    }
}
