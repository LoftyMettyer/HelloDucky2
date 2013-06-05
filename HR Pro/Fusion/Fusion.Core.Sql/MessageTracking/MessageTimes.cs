using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Fusion.Core.Sql
{
    public class MessageTimes
    {
        public DateTime? LastProcessedDate
        {
            get;
            set;
        }

        public DateTime? LastGeneratedDate
        {
            get;
            set;
        }
    }
}
