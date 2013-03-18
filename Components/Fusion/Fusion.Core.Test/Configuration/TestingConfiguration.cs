using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace Fusion.Core.Test
{
    public class TestingConfiguration : ITestingConfiguration
    {
        public TestingConfiguration()
        {
            this.MessagePath =  ConfigurationManager.AppSettings["MessagePath"];

            if (MessagePath == null)
            {
                MessagePath = ".";
            }

            this.Community = ConfigurationManager.AppSettings["Community"];
            if (Community == null)
            {
                Community = "test";
            }
        }
        public string MessagePath
        {
            get;
            private set;
        }

        public string Community
        {
            get;
            private set;
        }
    }
}
