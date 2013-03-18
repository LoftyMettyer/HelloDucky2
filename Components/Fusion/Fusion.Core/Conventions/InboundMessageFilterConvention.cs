// --------------------------------------------------------------------------------------------------------------------
// <copyright file="InboundMessageFilterConvention.cs" company="Advanced Health and Care Limited">
//   Copyright © 2011 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the inbound message filter convention class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Fusion.Core.Conventions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Fusion.Core.InboundFilters;
    using StructureMap.Configuration.DSL;
    using StructureMap.Graph;
  
    /// <summary>
    /// Registration convention for StructureMap. Checks a given type to see if it should be registered as an InboundFilterHandler.  If it implements both the generic IInboundFilterHandler and IInboundFilterHandler<T> 
    /// then just register under the more specific interface
    /// </summary>
    public class InboundMessageFilters : IRegistrationConvention
    {
        public void Process(Type type, Registry registry)
        {
            var interfaceTypes = FindClosingTypes(type, typeof(IInboundFilterHandler<>));

            if (interfaceTypes.Count() > 0)
            {
                foreach (var interfaceType in interfaceTypes)
                {
                    registry
                        .For(interfaceType)
                        .Add(type);
                }
            }
            else if (typeof(IInboundFilterHandler).IsAssignableFrom(type) && !type.IsGenericType)
            {
                registry
                    .For(typeof(IInboundFilterHandler))
                    .Add(type);
                    //.Named("Filters");
            }
        }

        private static IEnumerable<Type> FindClosingTypes(Type pluggedType, Type templateType)
        {
            if (pluggedType.IsAbstract || pluggedType.IsInterface) yield break;

            if (templateType.IsInterface)
            {
                foreach (var interfaceType in pluggedType.GetInterfaces().Where(type => type.IsGenericType && (type.GetGenericTypeDefinition() == templateType)))
                {
                    yield return interfaceType;
                }
            }
            else if (pluggedType.BaseType.IsGenericType && (pluggedType.BaseType.GetGenericTypeDefinition() == templateType))
            {
                yield return pluggedType.BaseType;
            }

            if (pluggedType.BaseType == typeof(object)) yield break;

            foreach (var interfaceType in FindClosingTypes(pluggedType.BaseType, templateType))
            {
                yield return interfaceType;
            }
        }
    }
}
