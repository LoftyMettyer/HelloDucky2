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

            //export.Connection ("sa", "asr", "openhr81pe", ".\\sql2014");
            export.Connection("sa", "asr", "npg_openhr8_2", "HARPDEV02");

            var result = export.ExportDefinition(4);

     //       Assert.Fail();
        }

        [TestMethod()]
        public void TestImport()
        {
            var import = new OpenHR.TestToLive.Repository();
            //import.Connection("sa", "asr", "openhr81pe", ".\\sql2014");
            import.Connection("sa", "asr", "npg_openhr8_2", "HARPDEV02");

            var result = import.ImportDefinitions();

            Assert.IsTrue(result == Enums.RepositoryStatus.DefinitionsImported);
        }

    }
}
