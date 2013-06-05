using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Fusion.Core.Sql
{

    /// <summary>
    /// A request to send a message of a particular type. This is an internal structure representing data coming from sql server service broker
    /// </summary>
    public class SendFusionMessageRequest
    {
        public string MessageType
        {
            get;
            set;
        }

        public string LocalId
        {
            get;
            set;
        }

        public DateTime TriggerDate
        {
            get;
            set;
        }

    }
}
