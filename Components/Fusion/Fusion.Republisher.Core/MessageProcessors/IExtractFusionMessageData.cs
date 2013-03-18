
namespace Fusion.Republisher.Core.MessageProcessors
{
    using System.Xml;

    public interface IExtractFusionMessageData
    {
        object ReadFromMessage(XmlNode message, XmlNamespaceManager ns, object currentState);

        void Update(XmlNode message, XmlNamespaceManager ns, object messageData);
    }
}
