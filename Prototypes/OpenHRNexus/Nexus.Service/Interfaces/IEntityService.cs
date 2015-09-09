using Nexus.Common.Enums;
using Nexus.Common.Models;
using System.Collections.Generic;

namespace Nexus.Service.Interfaces
{
    public interface IEntityService
    {
        IEnumerable<EntityModel> GetEntities(EntityType id);
    }
}
