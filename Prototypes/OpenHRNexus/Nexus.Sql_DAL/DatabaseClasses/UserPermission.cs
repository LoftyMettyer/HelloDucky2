using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Nexus.Repository.DatabaseClasses
{
	public class UserPermission
	{
		public IEnumerable<string> Roles;
		public IEnumerable<string> Claims;
		public IEnumerable<string> DataPermissions;

	}
}
