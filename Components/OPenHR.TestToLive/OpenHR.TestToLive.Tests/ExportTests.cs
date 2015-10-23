using OpenHR.TestToLive;
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenHR.TestToLive.Tests
{
    [TestClass]
    public class ExportTests
    {
        [TestMethod()]
        public void MainTest()
        {
            var repository = new Repository();
         //   Repository.Main();



            Assert.Fail();
        }

        [TestMethod()]
        public void TestExport()
        {
            var export = new OpenHR.TestToLive.Repository();

            export.Connection ("sa", "asr", "openhr81pe", ".\\sql2014");

            var result = export.ExportDefinition(44);

     //       Assert.Fail();
        }

        [TestMethod()]
        public void TestImport()
        {
            var import = new OpenHR.TestToLive.Repository();
            import.Connection("sa", "asr", "openhr81pe", ".\\sql2014");

            var result = import.ImportDefinitions();

            Assert.IsTrue(result == Enums.RepositoryStatus.DefinitionsImported);
        }

    }
}
