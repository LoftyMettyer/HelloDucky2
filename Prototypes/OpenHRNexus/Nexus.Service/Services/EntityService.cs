using System.Collections.Generic;
using Nexus.Common.Enums;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;
using Nexus.Common.Interfaces.Services;

namespace Nexus.Service.Services {
	public class EntityService : IEntityService {
		private IEntityRepository _entityRepository;


		public EntityService(IEntityRepository entityRepository) {
			_entityRepository = entityRepository;
		}

		public IEnumerable<EntityModel> GetEntities(EntityType id) {
			return _entityRepository.GetEntities(id);
		}
	}
}
