using System.Xml;
using Microsoft.Exchange.WebServices.Data;

namespace OutlookCalendarLogic.Exchange {
  class ExchangeTraceListener : ITraceListener {
	public void Trace(string traceType, string traceMessage) {
	  CreateXMLTextFile(traceType, traceMessage);
	}

	private void CreateXMLTextFile(string fileName, string traceContent) {
	  // Create a new XML file for the trace information.
	  try {
		// If the trace data is valid XML, create an XmlDocument object and save.
		XmlDocument xmlDoc = new XmlDocument();
		xmlDoc.Load(traceContent);
		xmlDoc.Save(fileName + ".xml");
	  } catch {
		// If the trace data is not valid XML, save it as a text document.
		System.IO.File.WriteAllText(fileName + ".txt", traceContent);
	  }
	}
  }
}
