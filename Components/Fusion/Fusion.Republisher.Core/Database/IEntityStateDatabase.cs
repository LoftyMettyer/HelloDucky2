
namespace Fusion.Republisher.Core.Database
{
    using System;
    using System.Collections.Generic;

    public interface IEntityStateDatabase
    {
        IEnumerable<Guid> GetAllEntityRefs(string community, string messageType);
        EntityState ReadEntityState(string community, string messageType, Guid entityRef);
        void UpdateEntityState(EntityState entityState);
    }
}
