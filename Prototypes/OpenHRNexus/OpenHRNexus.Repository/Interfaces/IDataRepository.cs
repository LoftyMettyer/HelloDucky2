using System.Collections.Generic;
using OpenHRNexus.Common.Models;

namespace OpenHRNexus.Repository.Interfaces
{
	public interface IDataRepository
	{
		IEnumerable<DynamicDataModel> GetData();
	}
}
