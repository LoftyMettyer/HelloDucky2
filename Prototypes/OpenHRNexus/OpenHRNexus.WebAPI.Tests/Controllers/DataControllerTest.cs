using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenHRNexus.Repository.SQLServer;
using OpenHRNexus.Service.Services;
using OpenHRNexus.WebAPI.Controllers;
using OpenHRNexus.Common.Models;

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


        [TestMethod]
        public void InstanciateProcess_IsNotNull()
        {
            var result = _mockController.InstantiateProcess(1, 1, false);
            Assert.IsNotNull(result);
        }

        [TestMethod]
        public void InstanciateProcess_ContainsWebformModels()
        {
            var result = _mockController.InstantiateProcess(1, 1, false);
            Assert.IsTrue(result is IEnumerable<WebFormModel>);
        }


    }


}
