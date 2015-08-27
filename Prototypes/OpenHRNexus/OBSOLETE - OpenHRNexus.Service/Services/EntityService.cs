using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Service.Interfaces;
using System.Collections.Generic;
using OpenHRNexus.Common.Enums;
using OpenHRNexus.Common.Models;

namespace OpenHRNexus.Service.Services
{
    public class EntityService : IEntityService
    {
        private IEntityRepository _entityRepository;


        public EntityService(IEntityRepository entityRepository)
        {
            _entityRepository = entityRepository;
        }

        public IEnumerable<EntityModel> GetEntities(EntityType? id)
        {
            return _entityRepository.GetEntities(id);
        }
    }
}
