// --------------------------------------------------------------------------------------------------------------------
// <copyright file="OutboundMessageFilterConvention.cs" company="Advanced Health and Care Limited">
//   Copyright © 2011 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the outbound message filter convention class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Fusion.Core.Conventions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using StructureMap.Configuration.DSL;
    using StructureMap.Graph;
    using Fusion.Core.OutboundFilters;

    /// <summary>
    /// Registration convention for StructureMap. Checks a given type to see if it should be registered as an OutboundFilterHandler.  If it implements both the generic IOutboundFilterHandler and IOutboundFilterHandler<T> 
    /// then just register under the more specific interface
    /// </summary>
    public class OutboundMessageFilters : IRegistrationConvention
    {
        public void Process(Type type, Registry registry)
        {
            var interfaceTypes = FindClosingTypes(type, typeof(IOutboundFilterHandler<>));

            if (interfaceTypes.Count() > 0)
            {
                foreach (var interfaceType in interfaceTypes)
                {
                    registry
                        .For(interfaceType)
                        .Add(type);
                }
            }
            else if (typeof(IOutboundFilterHandler).IsAssignableFrom(type) && !type.IsGenericType)
            {
                registry
                    .For(typeof(IOutboundFilterHandler))
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
