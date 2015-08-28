using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;
using System.Collections.Generic;

namespace Nexus.Common.Classes
{
    public class BusinessProcess : BaseEntity
    {
        //IDataRepository _repository;

        //public BusinessProcess(IDataRepository repository)
        //{
        //    _repository = repository;
        //}

        IEnumerable<BusinessProcessStep> Steps { get; set; }

        public WebForm GetFirstStep {
            get
            {
                return new WebForm();
            }
        }
    }

}

