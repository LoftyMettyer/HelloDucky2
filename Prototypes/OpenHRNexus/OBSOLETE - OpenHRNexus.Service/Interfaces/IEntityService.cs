using OpenHRNexus.Common.Enums;
using OpenHRNexus.Common.Models;
using System.Collections.Generic;

namespace OpenHRNexus.Service.Interfaces
{
    public interface IEntityService
    {
        IEnumerable<EntityModel> GetEntities(EntityType? id);
    }
}
