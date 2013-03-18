﻿using Fusion.Core.Sql;
using StructureMap.Configuration.DSL;

namespace ExampleConnector.Registries
{
    public class SendFusionMessageRequestBuilderRegistry : Registry
    {
        public SendFusionMessageRequestBuilderRegistry()
        {
            For<ISendFusionMessageRequestBuilder>().Use<SendFusionMessageRequestBuilder>();
        }
    }
}
