using Fusion.Republisher.Core.MessageProcessors;
using System;
using System.IO;
using System.Xml;

namespace Fusion.Republisher.Core
{
    public interface IMessageDefinition
    {
        FusionMessageDefinition[] MessageDefinition { get; }
        XmlNamespaceManager NamespaceManager { get; }
        Stream GetBlankXmlTemplate();
    }

    public interface IMessageValidator
    {
        IMessageValidatorResults ValidateMessage(string xml);     
    }

    public interface IMessageValidatorResults
    {
        bool IsValid
        {
            get;
        }

        string ValidationMessage
        {
            get;
        }
    }

    public class MessageValidationResults : IMessageValidatorResults
    {
        public bool IsValid
        {
            get;
            set;
        }

        public string ValidationMessage
        {
            get; set;
        }
    }
}