using Nexus.Common.Interfaces;
using System.Collections.Generic;

namespace Nexus.Common.Models
{
    public class WebFormField : WebFormControl
    {
        protected IDictionary _dictionary;
        string _title;

        public WebFormField() { }

        public WebFormField(IDictionary dictionary)
        {
            _dictionary = dictionary;
        }

        // TODO - Hacky bit - this should be ninjectable
        public void SetDictionary(IDictionary dictionary)
        {
            _dictionary = dictionary;
        }


        public int sequence { get; set; }
        public int columnid { get; set; }

        public string elementid { get
            {
                return string.Format("we_{0}_{1}", id.ToString(), sequence.ToString());
            }
        }

     //   Public DynamicColumn Column { get; set; }
        public string title {
            get
            {
                if (_dictionary != null)
                {
                    return _dictionary.GetTranslation(_title);
                }

                return _title;

            }
             set { _title = value; }
        }
        public string type { get; set; }
        public string value { get; set; }
        public bool required { get; set; }
        public bool disabled { get; set; }
        public List<WebFormFieldOption> options { get; set; }

    }
}
