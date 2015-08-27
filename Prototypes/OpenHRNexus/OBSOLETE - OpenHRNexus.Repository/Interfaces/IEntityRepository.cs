using OpenHRNexus.Common.Enums;
using OpenHRNexus.Common.Models;
using System.Collections.Generic;

namespace OpenHRNexus.Repository.Interfaces {
	public interface IEntityRepository {
		IEnumerable<EntityModel> GetEntities(EntityType? id);
	}

}
