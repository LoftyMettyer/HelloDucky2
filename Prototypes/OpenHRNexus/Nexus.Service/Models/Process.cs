using Nexus.Common.Classes;
using Nexus.Common.Interfaces;
using Nexus.Common.Models;
using System.Collections.Generic;
using System.Linq;

namespace Nexus.Service.Classes
{
    public class ProcessModel : Process
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

        IEnumerable<ProcessElement> Steps { get; set; }

    }

}

