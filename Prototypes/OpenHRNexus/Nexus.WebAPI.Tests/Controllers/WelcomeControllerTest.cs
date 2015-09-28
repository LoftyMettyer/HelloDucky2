using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Nexus.Common.Messages;
using Nexus.WebAPI.Controllers;
using Nexus.Common.Interfaces.Services;

namespace Nexus.WebAPI.Tests.Controllers {
	[TestClass]
	public class WelcomeControllerTest {
		[TestMethod]
		public void WelcomeController_Authenticate() {
			// Arrange
			var mockService = new Mock<IWelcomeService>();
			mockService.Setup(x => x.RequestAccount("SomeEmail", new System.Guid().ToString()));

			WelcomeController controller = new WelcomeController(mockService.Object);

			// Act
			RegisterNewUserMessage result = controller.Authenticate("UserName", new System.Guid().ToString());

			// Assert
			Assert.IsNotNull(result);
			//Assert.AreEqual(result.Role, "Employee");
		}

		[TestMethod]
		public void WelcomeController_GetRoles() {
			// Arrange
			var mockService = new Mock<IWelcomeService>();
			mockService.Setup(x => x.GetClaims(Guid.NewGuid()));

			WelcomeController controller = new WelcomeController(mockService.Object);

			var userId = Guid.NewGuid().ToString();
			var result = controller.GetClaims(userId);

			// Assert
			Assert.IsNotNull(result);
			//Assert.AreEqual(result.Role, "Employee");
		}



	}
}
