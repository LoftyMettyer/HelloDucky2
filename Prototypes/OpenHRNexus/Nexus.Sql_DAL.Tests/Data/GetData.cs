using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Repository.SQLServer;

namespace Nexus.Repository.Tests.Data
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

		[TestMethod]
		public void DataRepository_GetDataReturnsSingleRecord()
		{
			var dataRepository = new SqlDataRepository();
			var data = dataRepository.GetData(265);
			Assert.AreEqual(data.LongCount(), 1);
		}

		[TestMethod]
		public void DataRepository_GetDataReturnsMultipleRecords()
		{
			var dataRepository = new SqlDataRepository();
			var data = dataRepository.Data;
			Assert.IsTrue(data.LongCount() > 1, "The record count is greater than 1");

		}


	}

}
