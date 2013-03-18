using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.General;

namespace Fusion.Core
{
    public static class FusionUtils
    {
        public static string GetMessageName(this FusionMessage message) {

            string typeName = message.GetType().Name;

            if (typeName.EndsWith("Message"))
            {
                typeName = typeName.Substring(0, typeName.Length - 7);
            }
            if (typeName.EndsWith("Request"))
            {
                typeName = typeName.Substring(0, typeName.Length - 7);
            }

            return typeName;
        }
    }
}
