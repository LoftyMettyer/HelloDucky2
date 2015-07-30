using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using OpenHRNexus.Service.Interfaces;
using OpenHRNexus.WebAPI.Controllers;

namespace OpenHRNexus.WebAPI.Tests.Controllers
{
	[TestClass]
	public class DataControllerTest
	{

		[TestMethod]
		public void GetReportData_ReturnsNonNull()
		{
			// Arrange
			var mockService = new Mock<IDataService>();
			mockService.Setup(x => x.GetData(1));

			DataController controller = new DataController(mockService.Object);

			var result = controller.GetReportData(1.ToString());

			// Assert
			Assert.IsNotNull(result);
			//Assert.AreEqual(result.Role, "Employee");
		}

	}
}
