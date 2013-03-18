using System;
using NServiceBus;

namespace ExampleConnector
{
    public class WebStartup : IWantToRunAtStartup
    {
        public IBus Bus { get; set; }

        public void Run()
        {
            //Fusion.ConnectorWeb.WebStartup web = new Fusion.ConnectorWeb.WebStartup("http://localhost:80/Fusion/");
            //web.Go();
        }

        public void Stop()
        {

        }
    }
}