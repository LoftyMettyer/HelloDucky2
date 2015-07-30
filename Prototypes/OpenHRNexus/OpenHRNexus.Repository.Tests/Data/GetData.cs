using System;
using System.Text;
using System.Collections.Generic;
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

	}
}
