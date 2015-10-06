using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data.Entity;
using System.Linq;
using SystemManagerService;
using SystemManagerService.Enums;

namespace SystemManager.Tests
{
    [TestClass]
    public class PermissionTests //: TransactionTest
    {
        protected SecurityManager context;
        protected DbContextTransaction transaction;
        bool UseTransaction = false;

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
        public void AddPermissionGroup()
        {
            //var sysMan = new SystemManagerService.Structure();
          //  var secMan = new Roles();
            int originalRoleCount = context.PermissionGroups.Count();

            var message = context.AddPermissionGroup("SystemManager.Test", "An auto generated system test.");
            Assert.AreEqual(context.PermissionGroups.Count(), originalRoleCount + 1, "Incorrect amount of groups in list");
            Assert.IsTrue(message.ModifiedId > 0, "Valid Id not returned");

        //    return message.ModifiedId;

        }

        [TestMethod]
        public void AddPermissionGroup_NameUniqueCheck()
        {
            Assert.Fail("Group Name unique check not yet implemented");
        }


        [TestMethod]
        public void AddPermissionToGroup()
        {

            var groupID = context.PermissionGroups.FirstOrDefault().Id;
            var categoryId = 2;
            var facetId = 1;

            var message = context.AddPermissionToGroup(groupID, categoryId, facetId);
            Assert.AreEqual(message.status, SaveStatusEnum.Success, "Permission added failure");

        }

    }
}
