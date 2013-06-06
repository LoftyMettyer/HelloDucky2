using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.Xml;
using System.IO;

namespace Fusion.Publisher.SocialCare
{
    public class SerializationHelper
    {
        public static string ToXml<T>(T o) where T : class
        {
            XmlSerializer xmlOut = new XmlSerializer(typeof(T));

            StringBuilder stringBuilder = new StringBuilder();

            XmlWriterSettings writerSettings = new XmlWriterSettings();
            writerSettings.OmitXmlDeclaration = true;

            XmlWriter f = XmlWriter.Create(stringBuilder, writerSettings);

            xmlOut.Serialize(f, o);
            f.Close();

            return stringBuilder.ToString();
        }

        public static T FromXml<T>(string text) where T : class
        {
            StringReader stReader = new StringReader(text);
            XmlSerializer b = new XmlSerializer(typeof(T));
            object o = b.Deserialize(stReader);
            stReader.Close();

            return o as T;
        }
    }
}
