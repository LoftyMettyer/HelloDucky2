
namespace Fusion.Test
{
    using System;

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
    }
}
