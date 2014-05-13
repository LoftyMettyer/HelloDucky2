using System;
using System.Collections.Generic;
using System.Web;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Linq;

namespace ABS_Self_Service.Models
{
 
    [XmlRoot("widget")]
    public class widgetsModel
    {
        [XmlElement("widgetName")]
        public string widgetName { get; set; }

        [XmlElement("widgetPositionX")]
        public string widgetPositionX { get; set; }

        [XmlElement("widgetPositionY")]
        public string widgetPositionY { get; set; }

        [XmlElement("widgetSizeX")]
        public string widgetSizeX { get; set; }

        [XmlElement("widgetSizeY")]
        public string widgetSizeY { get; set; }

        [XmlElement("widgetDescription")]
        public string widgetDescription { get; set; }

        [XmlElement("widgetDisplayMode")]
        public string widgetDisplayMode { get; set; }

        [XmlElement("widgetUri")]
        public string widgetUri { get; set; }
       

        public List<widgetsModel> LoadModel(string fileName) 
        {
            List<widgetsModel> widgetList = new List<widgetsModel>();

            XDocument xmlDoc = XDocument.Load(fileName);

            var serializer = new XmlSerializer(typeof(widgetsModel));
            var model =
                from xml in xmlDoc.Descendants("widget")
                select serializer.Deserialize(xml.CreateReader()) as widgetsModel;

            foreach (widgetsModel widget in model) {
                widgetList.Add(widget);
            }

            return widgetList;
            
        }
    }

}

