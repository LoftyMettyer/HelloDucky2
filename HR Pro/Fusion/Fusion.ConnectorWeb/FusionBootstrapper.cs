using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Nancy.Bootstrappers.StructureMap;

namespace Fusion.Web
{
    public class FusionBootstrapper : StructureMapNancyBootstrapper
    {
        protected override void ApplicationStartup(StructureMap.IContainer container, Nancy.Bootstrapper.IPipelines pipelines)
        {
            base.ApplicationStartup(container, pipelines);


        }
    }
}
