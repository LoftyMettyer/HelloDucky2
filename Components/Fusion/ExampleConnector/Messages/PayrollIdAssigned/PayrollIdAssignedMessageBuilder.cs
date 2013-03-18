// --------------------------------------------------------------------------------------------------------------------
// <copyright file="PayrollIdAssignedMessageBuilder.cs" company="Advanced Health and Care Limited">
//   Copyright © 2011 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the payroll identifier assigned message builder class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace MyPublisher.OutboundBuilders
{
    using System;
    using Connector1.Configuration;
    using Connector1.DatabaseAccess;
    using Connector1.Messages;
    using Fusion.Core.Sql;
    using Fusion.Core.Sql.OutboundBuilder;
    using Fusion.Messages.Example;
    using Fusion.Messages.General;
        
    public class PayrollIdAssignedMessageBuilder : IOutboundBuilder
    {
        public PayrollIdAssignedMessageBuilder(IServiceUserDb serviceUserDb, IBusRefTranslator busRefTranslator, IFusionConfiguration config)
        {
            this.serviceUserDb = serviceUserDb;
            this.refTranslator = busRefTranslator;
            this.config = config;
        }

        private IServiceUserDb serviceUserDb;
        private IBusRefTranslator refTranslator;
        private IFusionConfiguration config;

        public FusionMessage Build(SendFusionMessageRequest source)
        {
            //var su = this.serviceUserDb.ReadServiceUser(Convert.ToInt32(source.LocalId));

            Guid busRef = this.refTranslator.GetBusRef(EntityTranslationNames.ServiceUser, source.LocalId);

            string xml = String.Format("<payrollIdAssigned><ref>{0}</ref><payrollId>{1}</payrollId></serviceUserUpdate>",
                busRef.ToString(), 1
                );

            return new PayrollIdAssignedMessage()
            {
                CreatedUtc = source.TriggerDate,
                Id = Guid.NewGuid(),
                Originator = config.ServiceName,
                EntityRef = busRef,
                Xml = xml
            };

        }
    }
}
