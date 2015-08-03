using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data.Entity;
using System.Linq;
using SystemManagerService;

namespace SystemManager.Tests
{
    [TestClass]
    public class RoleTests //: TransactionTest
    {
        protected SecurityManager context;
        protected DbContextTransaction transaction;
        bool UseTransaction = true;

        [TestInitialize]
        public void TestInitialize()
        {
            context = new SecurityManager();

            if (UseTransaction)
            {
                transaction = context.Database.BeginTransaction();
            }

        }

        [TestCleanup]
        public void TestCleanup()
        {
            if (UseTransaction)
            {
                transaction.Rollback();
                transaction.Dispose();
            }
            context.Dispose();
        }

        [TestMethod]
        public void CreateGroup()
        {
            //var sysMan = new SystemManagerService.Structure();
          //  var secMan = new Roles();
            int originalRoleCount = context.Groups.Count();

            context.AddRole("NewRole", "A description");
            Assert.AreEqual(context.Groups.Count(), originalRoleCount + 1);

        }

    }
}
