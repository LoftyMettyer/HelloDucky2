using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Fusion.Republisher.Core.Database
{
    /// <summary>
    /// A POCO representing republisher's view entity state
    /// </summary>
    public class EntityState
    {
        public string  Community;
        public string MessageType;
        public Guid? EntityRef;
        public Guid? PrimaryEntityRef;
        public DateTime LastUpdate;
        public string MessageState;
    }
}
