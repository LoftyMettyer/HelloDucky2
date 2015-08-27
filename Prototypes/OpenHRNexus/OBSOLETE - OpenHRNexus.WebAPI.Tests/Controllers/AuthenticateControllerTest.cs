using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using OpenHRNexus.Repository.Messages;
using OpenHRNexus.Service.Interfaces;
using OpenHRNexus.WebAPI.Controllers;

namespace OpenHRNexus.WebAPI.Tests.Controllers {
	[TestClass]
	public class AuthenticateControllerTest {
		[TestMethod]
		public void Authenticate() {
			// Arrange
			var mockService = new Mock<IAuthenticateService>();
			mockService.Setup(x => x.RequestAccount("SomeEmail", new System.Guid().ToString()));

			AuthenticateController controller = new AuthenticateController(mockService.Object);

			// Act
			RegisterNewUserMessage result = controller.Authenticate("UserName", new System.Guid().ToString());

			// Assert
			Assert.IsNotNull(result);
			//Assert.AreEqual(result.Role, "Employee");
		}

		[TestMethod]
		public void GetRoles() {
			// Arrange
			var mockService = new Mock<IAuthenticateService>();
			mockService.Setup(x => x.GetClaims(Guid.NewGuid()));

			AuthenticateController controller = new AuthenticateController(mockService.Object);

			var userId = Guid.NewGuid().ToString();
			var result = controller.GetClaims(userId);

			// Assert
			Assert.IsNotNull(result);
			//Assert.AreEqual(result.Role, "Employee");
		}



	}
}
