using Nexus.Common.Interfaces.Repository;
using System;
using Nexus.Common.Interfaces;
using Nexus.Common.Enums;
using Nexus.Common.Models;

namespace Nexus.Sql_Repository
{
    public class SqlDesignRepository : IDesignRepository
    {
        public DesignStatus AddColumnToTable(string name, TableModel table, IEntity column)
        {
            throw new NotImplementedException();
        }

        public DesignStatus AddTable(TableModel table)
        {

            // Populate the metadata dictionary

            // Checks and stuff for usage

            // Validation - this will actually work

            // Commit the change

            // Return a bit more than just ok - perhaps the generated object class?
            return DesignStatus.Success;

        }

    }
}
