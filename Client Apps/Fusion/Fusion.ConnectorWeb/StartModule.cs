using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Nancy;
using Fusion.ConnectorWeb;

namespace Fusion.Web
{
    public class StartModule : NancyModule
    {
        public StartModule()
        {
            Get["/"] = _ => "Fusion Connector - Hello World!";
            Get["/messages"] = _ => String.Join("<p>", (new ConfigureMemoryAppender()).ReadTest());
        }
    }
}
