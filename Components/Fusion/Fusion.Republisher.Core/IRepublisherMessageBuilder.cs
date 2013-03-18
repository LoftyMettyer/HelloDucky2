using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.General;

namespace Fusion.Republisher.Core
{
    public interface IRepublisherMessageBuilder
    {
        FusionMessage Build(PublishMessageRequest source);        
    }
}
