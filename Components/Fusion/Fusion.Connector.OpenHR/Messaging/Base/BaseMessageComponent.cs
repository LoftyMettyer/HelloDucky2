using System.ComponentModel;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Fusion.Connector.OpenHR.Configuration;
using StructureMap.Attributes;

namespace Fusion.Connector.OpenHR.Messaging.Base
{
	public class BaseMessageComponent
	{

		private int _version = 1;

		[SetterProperty]
		public static IFusionConfiguration config { get; set; }

		public string ToXml()
		{
			var xsSubmit = new XmlSerializer(GetType());
			var sww = new StringWriter();
			var writer = XmlWriter.Create(sww);
			xsSubmit.Serialize(writer, this);
			return sww.ToString();
		}


		[XmlAttribute]
		public int version
		{
			get
			{
				return _version;
			}
			set
			{
				_version = value;
			}
		}

	}
}
