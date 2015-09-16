using Nexus.Common.Interfaces;
using Nexus.Common.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Nexus.Common.Classes
{
    public class ProcessFormElement
    {
        //IDictionary _dictionary;

    //    private WebForm() { }

        //public WebForm(IDictionary dictionary) : base()
        //{
        //    _dictionary = dictionary;
        //}


        public int id { get; set; }
        public string Name { get; set; }
        public List<WebFormField> Fields { get; set; }
        public List<WebFormButton> Buttons { get; set; }

        public void SetButtonEndpoints(Guid stepId)
        {
            foreach (var button in Buttons)
            {
                button.TargetUrl = string.Format("{0}/stepId={1}", button.TargetUrl, stepId);
            }
        }

        /// <summary>
        /// Translate all the display entities given a specified dictionary
        /// At present each element is called with a separate round trip to the server. Google says that entity framework will be 
        /// caching stuff, however some load testing and profiling will find out for sure. This function may need refractoring.
        /// </summary>
        /// <param name="dictionary"></param>
        public void Translate(IDictionary dictionary)
        {

            Name = dictionary.GetTranslation(Name);

            foreach (var field in Fields)
            {
                field.title = dictionary.GetTranslation(field.title);
            }

            foreach (var button in Buttons)
            {
                button.Title = dictionary.GetTranslation(button.Title);
            }
         
            foreach (var lookup in Fields.Where(f => f.IsLookupValue))
            {
                lookup.options = dictionary.GetLookupValues(lookup.columnid);
            }



        }

    }

}
