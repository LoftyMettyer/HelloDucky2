﻿using System.Collections.Generic;

namespace Nexus.Sql_Repository.DatabaseClasses.Security
{
	public class UserPermission {
		public IEnumerable<string> Roles;
		public IEnumerable<string> Claims;
		public IEnumerable<string> DataPermissions;

	}
}