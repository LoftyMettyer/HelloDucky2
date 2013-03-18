using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fusion.Republisher.Core.MessageProcessors
{
    public class FusionMessageDefinition
    {
        public string Tag;
        public IExtractFusionMessageData MessageExtractor;
    }
}
