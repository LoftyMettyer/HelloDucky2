﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenHRTestToLive.Enums;

namespace OpenHRTestToLive.Tests
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
            var export = new OpenHRTestToLive.Repository();
            var outputFilename = "c:\\dev\\absdef1.xml";

      export.Connection ("sa", "asr", "openhr81pe", ".\\sql2014");
      //export.Connection("sa", "asr", "npg_openhr8_2", "HARPDEV02");

      var result = export.ExportDefinition(10, outputFilename);

            Assert.IsInstanceOfType(result, typeof(string));
        }

        [TestMethod()]
        public void TestImport()
        {
            var import = new Repository();
            import.Connection("sa", "asr", "func8.2", ".\\sql2014");

            //var inputFile = "c:\\dev\\absdef1.xml";
            var inputFile = "C:\\WorkflowExports\\SimpleCalc.xml";

            var result = import.ImportDefinitions(inputFile);

            Assert.IsTrue(result == RepositoryStatus.DefinitionsImported);
        }


    [TestMethod()]
    public void TestImportInvalidFile()
      {

        var import = new Repository();
        import.Connection("sa", "asr", "openhr81pe", ".\\sql2014");

        const string inputFile = "c:\\dev\\junkfile2.xml";

        var result = import.ImportDefinitions(inputFile);

        Assert.IsTrue(result == RepositoryStatus.Error);

    }

    }
}
