using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Repository.SQLServer;
using Nexus.Service.Services;
using Nexus.WebAPI.Controllers;
using Nexus.Common.Models;
using Moq;
using Nexus.Service.Interfaces;

namespace Nexus.WebAPI.Tests.Controllers
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
        public void InstantiateProcess_IsNotNull()
        {
            var result = _mockController.InstantiateProcess(1, 1, false);
            Assert.IsNotNull(result);
        }

        [TestMethod]
        public void InstantiateProcess_ContainsWebformModels()
        {
            var result = _mockController.InstantiateProcess(1, 15, false);
            Assert.IsTrue(result is IEnumerable<WebFormModel>);
        }

        [TestMethod]
        public void InstantiateProcess_ContainsSingleWebformModel()
        {

            //var mockService = new Mock<IDataService>();
            //mockService.Setup(x => x.GetClaims(Guid.NewGuid()));

            //AuthenticateController controller = new AuthenticateController(mockService.Object);

            //var userId = Guid.NewGuid().ToString();
            //var result = controller.GetClaims(userId);





            var result = (List <WebFormModel>)_mockController.InstantiateProcess(1, 16, false);
            Assert.IsTrue(result.Count == 1);
        }

    }


}
