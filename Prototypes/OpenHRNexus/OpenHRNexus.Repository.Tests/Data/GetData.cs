using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenHRNexus.Repository.SQLServer;

namespace OpenHRNexus.Repository.Tests.Data
{
	[TestClass]
	public class GetData
	{
		[TestMethod]
		public void DataRepository_IsNotNull()
		{
			var dataRepository = new SqlDataRepository();
			Assert.IsNotNull(dataRepository);
		}

		[TestMethod]
		public void DataRepository_GetDataReturnsEmptyCollectionForInvalidId()
		{
			var dataRepository = new SqlDataRepository();
			var data = dataRepository.GetData(0);
			Assert.IsNotNull(data);
			Assert.AreEqual(data.LongCount(), 0);

		}

	}

}
