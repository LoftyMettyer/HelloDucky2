using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters;
using System.Text;

namespace Fusion.Republisher.Core.MessageStateSerializer
{
    public class JsonMessageStateSerializer : IMessageStateSerializer
    {
        public MessagePersistedState Deserialize(string messageState)
        {
            if (String.IsNullOrEmpty(messageState))
            {
                return new MessagePersistedState
                {
                    State = new Dictionary<string,object>()
                };
            }

            var js = CreateSerializer();
            
            StringReader sr = new StringReader(messageState);
            var messageData = js.Deserialize<MessagePersistedState>(new JsonTextReader(sr));

            return messageData;
        }

        public string Serialize(MessagePersistedState state)
        {
            var js = CreateSerializer();

            StringWriter sw = new StringWriter();
            js.Serialize(sw, state);

            return sw.ToString();
        }

        private JsonSerializer CreateSerializer()
        {
            var serializerSettings = new JsonSerializerSettings
            {
                TypeNameAssemblyFormat = FormatterAssemblyStyle.Simple,
                TypeNameHandling = TypeNameHandling.Auto,
                Converters = { new IsoDateTimeConverter { DateTimeStyles = DateTimeStyles.RoundtripKind } }
            };

            JsonSerializer js = JsonSerializer.Create(serializerSettings);

            return js;
        }
    }
}
