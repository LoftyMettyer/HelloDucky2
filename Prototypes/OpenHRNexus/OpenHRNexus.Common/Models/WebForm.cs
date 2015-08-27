using OpenHRNexus.Common.Interfaces;
using System.Collections.Generic;

namespace OpenHRNexus.Common.Models
{
    public class WebForm : ITranslate
    {
        public int id { get; set; }
        public string Name { get; set; }
        public List<WebFormField> Fields { get; set; }

        public void Translate(string language)
        {
            foreach (WebFormField field in Fields)
            {
                if (language == "FR-FR" && field.field_type == "textfield")
                {
                    field.field_value = "La " + field.field_value;
                }

            }
        }

    }

}
