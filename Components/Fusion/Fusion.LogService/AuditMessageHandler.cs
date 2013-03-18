using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Messaging;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters;
using Fusion.Messages.General;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using NServiceBus;
using NServiceBus.Unicast.Queuing.Msmq;
using NServiceBus.Unicast.Transport;
using NServiceBus.Unicast.Transport.Transactional;
using NServiceBus.Utils;
using Fusion.LogService.DatabaseAccess;
using StructureMap.Attributes;

namespace Fusion.LogService
{
    class Boostrapper : IWantToRunAtStartup
    {
        public IBus Bus { get; set; }
        public ITransport AuditMessageQueueTransport;
        public TransactionalTransport CurrentEndpointTransport { get; set; }

        [SetterProperty]
        public ILogDatabase LogDatabase
        {
            get;
            set;
        }

        public void Run()
        {
            // Create an in-memory TransactionalTransport and point it to the AuditQ
            // Serialize data into the database.

            string auditQueue =  ConfigurationManager.AppSettings["AuditQueue"];
            string machineName = Environment.MachineName;

            // Make sure that the queue being monitored, exists!
            if (!MessageQueue.Exists(MsmqUtilities.GetFullPathWithoutPrefix(auditQueue)))
            {
                // The error queue being monitored must be local to this endpoint
                throw new Exception(string.Format("The audit queue {0} being monitored must be local to this endpoint and must exist. Make sure a transactional queue by the specified name exists. The audit queue to be monitored is specified in the app.config", auditQueue));
            }

            // Create an in-memory transport with the same configuration as that of the current endpoint.
            AuditMessageQueueTransport = new TransactionalTransport()
            {
                IsTransactional = CurrentEndpointTransport.IsTransactional,
                MaxRetries = CurrentEndpointTransport.MaxRetries,
                IsolationLevel = CurrentEndpointTransport.IsolationLevel,
                MessageReceiver = new MsmqMessageReceiver(),
                NumberOfWorkerThreads = CurrentEndpointTransport.NumberOfWorkerThreads,
                TransactionTimeout = CurrentEndpointTransport.TransactionTimeout,
                FailureManager = CurrentEndpointTransport.FailureManager
            };

            AuditMessageQueueTransport.TransportMessageReceived += new EventHandler<TransportMessageReceivedEventArgs>(AuditMessageQueueTransport_TransportMessageReceived);
            AuditMessageQueueTransport.Start(new Address(auditQueue, machineName));
        }

        public void Stop()
        {
            AuditMessageQueueTransport.TransportMessageReceived -= new EventHandler<TransportMessageReceivedEventArgs>(AuditMessageQueueTransport_TransportMessageReceived);
        }

        void AuditMessageQueueTransport_TransportMessageReceived(object sender, TransportMessageReceivedEventArgs e)
        {
            var message = e.Message;

            var serializerSettings = new JsonSerializerSettings
            {
                TypeNameAssemblyFormat = FormatterAssemblyStyle.Simple,
                TypeNameHandling = TypeNameHandling.Auto,
                Converters = { new IsoDateTimeConverter { DateTimeStyles = DateTimeStyles.RoundtripKind } }
            };

            JsonSerializer js = JsonSerializer.Create(serializerSettings);

            js.Binder = new FusionMessageSerializationBinder();
            var sr = new StreamReader(new MemoryStream(message.Body), true);

            var m = js.Deserialize<object[]>(new JsonTextReader(sr));

            if (m.Length != 1)
                return;

            FusionMessage msg = m[0] as FusionMessage;
            if (msg == null)
                return; 
            
            // Get the header list as a key value dictionary...
            Dictionary<string, string> headerDictionary = message.Headers.ToDictionary(k => k.Key, v => v.Value);

            var enclosedMessageType = headerDictionary["NServiceBus.EnclosedMessageTypes"];
            string messageType = enclosedMessageType;
            if (!String.IsNullOrWhiteSpace(enclosedMessageType))
            {
                messageType = enclosedMessageType.Split(new char[] { ',' }, StringSplitOptions.None)[0];
            }

            //AuditMessage messageToStore = new AuditMessage
            //{
            //    MessageId = message.Id,
            //    OriginalMessageId = message.GetOriginalId(),
            //    Body = messageBodyXml,
            //    Headers = headerDictionary,
            //    MessageType = messageType,
            //    ReceivedTime = DateTime.ParseExact(headerDictionary["NServiceBus.TimeSent"], "yyyy-MM-dd HH:mm:ss:ffffff Z", System.Globalization.CultureInfo.InvariantCulture)
            //};

            LogDatabase.AddMessageAudit(
                message.Id, message.GetOriginalId(), "", messageType, DateTime.ParseExact(headerDictionary["NServiceBus.TimeSent"], "yyyy-MM-dd HH:mm:ss:ffffff Z", System.Globalization.CultureInfo.InvariantCulture), message.ReplyToAddress.ToString(),
                msg.Id, msg.Originator, msg.EntityRef, msg.PrimaryEntityRef, msg.Community, msg.Xml, msg.CreatedUtc, msg.SchemaVersion);
            
            //Persister.Persist(messageToStore);

        }

    }

    internal class FusionMessageSerializationBinder : SerializationBinder
    {
        
        public override Type BindToType(string assemblyName, string typeName)
        {
            return typeof(ConcreteFusionMessage);
        }

       
    }

    internal class ConcreteFusionMessage : FusionMessage { }
}