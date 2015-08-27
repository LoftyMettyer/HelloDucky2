using Nexus.Common.Classes;
using Nexus.Common.Interfaces;
using Nexus.Common.Models;
using System.Collections.Generic;
using System.Linq;

namespace Nexus.Service.Classes
{
    public class BusinessProcessModel : BusinessProcess
    {
        //private IDataRepository _repository;

        //public BusinessProcessModel(IDataRepository repository)
        //{
        //    _repository = repository;
        //}


        //public void Populate(int id)
        //{
        //    this = _repository.GetBusinessProcess(id);
        //}

        IEnumerable<BusinessProcessStep> Steps { get; set; }

        public WebForm GetFirstStep()
        {

            var blah = Steps.Where(w => w.Id == 123).First();

            //var webForm = WebForms.Where(w => w.id == id).First();

            //var webForm2 = Steps

            //// TODO - Need these 2 because the above is not loading on demand. I'm sure there's some linq that does this, but off the top of my head I don't know what it is.
            //List<WebFormField> fields = WebFormFields.ToList();
            //List<WebFormFieldOption> options = WebFormFieldOptions.ToList();

            //return webForm;


            return new WebForm();
        }

    }

}

