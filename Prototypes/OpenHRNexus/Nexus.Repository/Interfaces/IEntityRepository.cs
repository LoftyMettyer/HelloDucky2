using Nexus.Common.Enums;
using Nexus.Common.Models;
using System.Collections.Generic;

namespace Nexus.Repository.Interfaces {
	public interface IEntityRepository {
		IEnumerable<EntityModel> GetEntities(EntityType? id);
	}

}
