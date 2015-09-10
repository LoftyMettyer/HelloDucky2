using Nexus.Common.Interfaces;
using System.Collections.Generic;

namespace Nexus.Common.Models
{
    public class WebForm
    {
        //IDictionary _dictionary;

        private WebForm() { }

        //public WebForm(IDictionary dictionary) : base()
        //{
        //    _dictionary = dictionary;
        //}


        public int id { get; set; }
        public string Name { get; set; }
        public List<WebFormField> Fields { get; set; }
        public List<WebFormButton> Buttons { get; set; }

    }

}
