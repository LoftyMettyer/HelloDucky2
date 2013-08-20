using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace RCVS.Helpers
{
	public class XMLHelper
	{
		public string SerializeToXml<T>(T value)
		{
			if (value == null)
			{
				return null;
			}

			//We need to override the root attribute of the XML string to be "Parameters", as required by the web services
			XmlAttributes attributes = new XmlAttributes();
			attributes.XmlRoot = new XmlRootAttribute("Parameters");
			XmlAttributeOverrides overrides = new XmlAttributeOverrides();
			overrides.Add(typeof(T), attributes);
			var serializer = new XmlSerializer(typeof(T), overrides);

			XmlWriterSettings settings = new XmlWriterSettings
			{
				Encoding = new UnicodeEncoding(false, false),
				Indent = false,
				OmitXmlDeclaration = true,
				ConformanceLevel = ConformanceLevel.Auto
			};

			using (StringWriter textWriter = new StringWriter())
			{
				using (XmlWriter xmlWriter = XmlWriter.Create(textWriter, settings))
				{
					serializer.Serialize(xmlWriter, value);
				}
				return textWriter.ToString();
			}
		}

		public T DeserializeFromXmlToObject<T>(string xml)
		{
			if (string.IsNullOrEmpty(xml))
			{
				return default(T);
			}

			var serializer = new XmlSerializer(typeof(T));

			XmlReaderSettings settings = new XmlReaderSettings();

			using (StringReader textReader = new StringReader(xml))
			{
				using (XmlReader xmlReader = XmlReader.Create(textReader, settings))
				{
					return (T)serializer.Deserialize(xmlReader);
				}
			}
		}
	}
}
