using Nexus.Common.Interfaces;
using System.Collections.Generic;

namespace Nexus.Common.Models
{
    public class WebForm : ITranslate //, IDataRepository
    {
        public int id { get; set; }
        public string Name { get; set; }
        public List<WebFormField> Fields { get; set; }
        public List<WebFormButton> Buttons { get; set; }

        public void Translate(string language)
        {
            foreach (WebFormField field in Fields)
            {
                if (language == "FR-FR" && field.type == "textfield")
                {
                    field.value = "La " + field.value;
                }
            }
        }



    }

}
