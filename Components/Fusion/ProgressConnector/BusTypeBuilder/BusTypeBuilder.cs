namespace ProgressConnector.BusTypeBuilder
{
    using Fusion.Messages.General;
    using ProgressConnector.ProgressInterface;
    using StructureMap;

    public class BusTypeBuilder : IBusTypeBuilder
    {

        public FusionMessage Build(OpenExchangeMessage source)
        {
            var fusionMessage = ObjectFactory.GetNamedInstance<FusionMessage>(source.MessageType + "FusionMessage");

            fusionMessage.CreatedUtc = source.Created;
            fusionMessage.Id = source.Id;
            fusionMessage.Originator = source.Originator;
            fusionMessage.EntityRef = source.EntityRef;
            fusionMessage.Xml = source.Xml;
            fusionMessage.SchemaVersion = source.SchemaVersion;
            fusionMessage.Community = source.Community;

            return fusionMessage;
        }
    }
}
