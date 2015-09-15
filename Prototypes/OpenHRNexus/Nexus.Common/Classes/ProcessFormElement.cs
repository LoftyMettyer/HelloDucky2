using Nexus.Common.Models;
using System;
using System.Collections.Generic;

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
                button.targeturl = string.Format("{0}/stepId={1}", button.targeturl, stepId);
            }
        }

    }

}
