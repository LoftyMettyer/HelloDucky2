using Connector1.ProgressInterface;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Fusion.Messages.General;

namespace Test.ProgressConnector
{
    
    
    /// <summary>
    ///This is a test class for OpenExchangeFusionMessageConvertorTest and is intended
    ///to contain all OpenExchangeFusionMessageConvertorTest Unit Tests
    ///</summary>
    [TestClass()]
    public class OpenExchangeFusionMessageConvertorTest
    {


        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        public class TestFusionMessage : FusionMessage
        {

        }

        /// <summary>
        ///A test for BuildOpenExchangeMessage
        ///</summary>
        [TestMethod()]
        public void BuildOpenExchangeMessageTest()
        {
            OpenExchangeFusionMessageConvertor target = new OpenExchangeFusionMessageConvertor(); // TODO: Initialize to an appropriate value
            FusionMessage message = new TestFusionMessage()
            {
                Id = new Guid("870A18AD-9D47-4D3D-AF0E-2598A77927BD"),
                CreatedUtc = new DateTime(2012, 12, 25, 14, 15, 23, 111, DateTimeKind.Utc),
                Originator = "me",
                SchemaVersion = 2,
                Xml = "<abc></abc>"
            };

            var actual = target.BuildOpenExchangeMessage(message);
            

        }
    }
}
