using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Fusion.Core.Sql
{
    public static class BusRefTranslatorExtensions
    {
        /// <summary>
        /// Gets the bus reference for a local id. This will create a bus reference if one does not exist already
        /// </summary>
        /// <param name="translationName"> Name of the translation. </param>
        /// <param name="localId">         Identifier for the local. </param>
        /// <returns>
        /// The bus reference.
        /// </returns>
        public static Guid GetBusRef(this IBusRefTranslator busTranslator, string translationName, string localId)
        {
            var busRef = busTranslator.TryGetBusRef(translationName, localId, true);
            return busRef.Value;
        }
    }
}