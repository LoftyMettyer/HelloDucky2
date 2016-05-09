using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenHRWorkflow.UnitTests
{
   [TestClass]
   public class OpenAmRestCallsTests
   {
      [TestMethod]
      public void LoginWithIncorrectCredentials()
      {
         var loginWithIncorrectCredentials = OpenAmRestCalls.LoginAndReturnToken("harry.combrink", "WrongPassword");
         Assert.AreEqual(loginWithIncorrectCredentials, null);
      }

      [TestMethod]
      public void LoginWithCorrectCredentials()
      {
         var loginWithCorrectCredentials = OpenAmRestCalls.LoginAndReturnToken("harry.combrink", "harryharry1");
         Assert.AreNotEqual(loginWithCorrectCredentials, null);
      }

      [TestMethod]
      public void GetUserIdFromGetIdFromSession()
      {
         var loginWithCorrectCredentials = OpenAmRestCalls.LoginAndReturnToken("harry.combrink", "harryharry1");
         var workspaceUserId = OpenAmRestCalls.GetIdFromSession(loginWithCorrectCredentials);
         Assert.AreNotEqual(workspaceUserId, null);
      }
   }
}
