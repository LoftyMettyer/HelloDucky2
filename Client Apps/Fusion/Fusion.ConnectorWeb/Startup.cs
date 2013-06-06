using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Web;

namespace Fusion.ConnectorWeb
{
    public class WebStartup
    {
        public WebStartup(string url)
        {
            this.url = url;
        }

        private string url;

        public void Go()
        {
            // Start NANCY listening

            Nancy.Hosting.Self.NancyHost host = new Nancy.Hosting.Self.NancyHost(
                new FusionBootstrapper(),
                new Uri(url));

            host.Start();

            // Configure in-memory logger

            ConfigureMemoryAppender cf = new ConfigureMemoryAppender();
            cf.Start();
        }
    }
}
