using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using OpenHRNexus.Repository.SQLServer;
using OpenHRNexus.Service.Interfaces;
using OpenHRNexus.Service.Services;
using OpenHRNexus.WebAPI.Controllers;

namespace OpenHRNexus.WebAPI.Tests.Controllers
{
	[TestClass]
	public class DataControllerTest
	{

        SqlDataRepository _mockRepository;
        DataService _mockService;
        DataController _mockController;

        [TestInitialize]
		public void TestInitialize()
		{
            _mockRepository = new SqlDataRepository();
            _mockService = new DataService(_mockRepository);
            _mockController = new DataController(_mockService);
        }


        [TestMethod]
		public void GetReportData_ReturnsNonNullForSingleRow()
		{
			// Arrange
	//		var mockService = new Mock<IDataService>();
//			mockService.Setup(x => x.GetData(78));

			var result = _mockController.GetReportData(78.ToString());
			Assert.IsNotNull(result);
		}

		[TestMethod]
		public void GetReportData_ReturnsNonNullForMultipleRows()
		{
			// Arrange
//			var mockService = new Mock<IDataService>();
//			mockService.Setup(x => x.GetData());
//			DataController controller = new DataController(mockService.Object);

			var result = _mockController.GetReportData("nothing");
			Assert.IsNotNull(result);
		}



	}


}
