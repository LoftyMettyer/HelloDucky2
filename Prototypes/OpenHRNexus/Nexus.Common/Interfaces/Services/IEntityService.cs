using Nexus.Common.Enums;
using Nexus.Common.Models;
using System.Collections.Generic;

namespace Nexus.Common.Interfaces.Services
{
    public interface IEntityService
    {
        IEnumerable<EntityModel> GetEntities(EntityType id);
    }
}
