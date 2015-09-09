using System;
using Nexus.Common.Interfaces.Services;
using Nexus.Common.Interfaces.Repository;
using Nexus.Sql_Repository.DatabaseClasses.Structure;
using Nexus.Common.Interfaces;
using Nexus.Common.Models;
using Nexus.Common.Enums;

namespace Nexus.Service.Services
{
    public class DesignerService : IDesignerService
    {
        private IDesignRepository _designRepository;

        public DesignerService(IDesignRepository designRepository)
        {
            _designRepository = designRepository;
        }


        public DesignStatus AddTable(string name)
        {
            var newTable = new TableModel() { Name = name };

            return _designRepository.AddTable(newTable);
        }
    }
}
