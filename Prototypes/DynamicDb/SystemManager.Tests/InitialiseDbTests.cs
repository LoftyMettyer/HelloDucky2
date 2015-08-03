using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using SystemManagerService;

namespace SystemManager.Tests
{
    [TestClass]
    public class InitialiseDbTests
    {
        [TestMethod]
        public void InitialiseDb_Categories()
        {
             var secMan = new SecurityManager();
            int originalRoleCount = secMan.PermissionCategories.Count();
            Assert.AreEqual(38, originalRoleCount);
        }

    }
}
