using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ProgressConnector.ProgressInterface
{
    public class ProgressSendStatus
    {
        public bool Error
        {
            get;
            set;
        }

        public string ErrorText
        {
            get;
            set;
        }
    }
}
