// --------------------------------------------------------------------------------------------------------------------
// <copyright file="IBusRefTranslator.cs" company="Advanced Health and Care Limited">
//   Copyright © 2011 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Declares the IBusRefTranslator interface
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Fusion.Core.Sql
{
    using System;

    /// <summary>
    /// Interface for bus reference translator. This provides a simple way to translate between bus references (guid) and local references
    /// </summary>
    public interface IBusRefTranslator
    {
        

        /// <summary>
        /// Gets a local reference number for a given bus reference
        /// </summary>
        /// <param name="translationName"> Name of the translation. </param>
        /// <param name="busRef">          The bus reference. </param>
        /// <returns>
        /// The local reference.
        /// </returns>
        string GetLocalRef(string translationName, Guid busRef);

        /// <summary>
        /// Sets the bus reference associated with a given local id
        /// </summary>
        /// <param name="translationName"> Name of the translation. </param>
        /// <param name="localId">         Identifier for the local. </param>
        /// <param name="busRef">          The bus reference. </param>
        void SetBusRef(string translationName, string localId, Guid busRef);

        /// <summary>
        /// Try to get bus reference. If the bus reference does not exist, indicate this
        /// </summary>
        /// <param name="translationName"> Name of the translation. </param>
        /// <param name="localId">         Identifier for the local. </param>
        /// <param name="canCreate">       (optional) the can create. </param>
        /// <returns>
        /// .
        /// </returns>
        BusTranslationResults TryGetBusRef(string translationName, string localId, bool canCreate = false);
    }
}
