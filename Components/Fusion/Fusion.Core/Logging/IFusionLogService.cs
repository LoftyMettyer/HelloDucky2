
namespace Fusion.Core.Logging
{
    using System;
using Fusion.Messages.General;

    public interface IFusionLogService
    {

        /// <summary>
        /// Logs a reference translation has been made. This log message will form part of the current transaction.
        /// </summary>
        /// <param name="entityName"> Name of the entity. </param>
        /// <param name="busRef">     The bus reference. </param>
        /// <param name="localId">    Identifier for the local. </param>
        void LogRefTranslationTransactional(string community, string entityName, Guid busRef, string localId);


        /// <summary>
        /// Logs an information message. This log message will form part of the current transaction and will always be sent
        /// </summary>
        /// <param name="message"> The message. </param>
        void InfoMessageTransactional(string community, Guid messageId, FusionLogLevel logLevel, Guid? entityRef, Guid? primaryEntityRef, string messageName, string message);


        /// <summary>
        /// Logs an information message.  This message will NOT form part of the current transaction and will always be sent
        /// </summary>
        /// <param name="message"> The message. </param>
        void InfoMessageNonTransactional(string community, Guid messageId, FusionLogLevel logLevel, Guid? entityRef, Guid? primaryEntityRef, string messageName, string message);

    }
}
