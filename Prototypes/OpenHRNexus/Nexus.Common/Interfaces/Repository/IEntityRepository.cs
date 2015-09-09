using System.Collections.Generic;
using Nexus.Common.Enums;
using Nexus.Common.Models;

namespace Nexus.Common.Interfaces.Repository {
	public interface IEntityRepository {
		IEnumerable<EntityModel> GetEntities(EntityType id);
	}

}
