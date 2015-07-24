using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using OpenHRNexus.Interfaces.Common;
using OpenHRNexus.Service.Interfaces;
using OpenHRNexus.WebAPI.Controllers;

namespace OpenHRNexus.WebAPI.Tests.Controllers {
	[TestClass]
	public class AuthenticateControllerTest {
		[TestMethod]
		public void Authenticate() {
			// Arrange
			var mockService = new Mock<IAuthenticateService>();
			mockService.Setup(x => x.RequestAccount("SomeEmail"));

			AuthenticateController controller = new AuthenticateController(mockService.Object);

			// Act
			INexusUser result = controller.Authenticate("UserName");

			// Assert
			Assert.IsNotNull(result);
			Assert.AreEqual(result.Role, "Employee");
		}
	}
}
