using Nexus.Common.Enums;
using Nexus.Common.Models;

namespace Nexus.Common.Interfaces.Repository
{
    public interface IDesignRepository
    {
        DesignStatus AddTable(TableModel table);
        DesignStatus AddColumnToTable(string name, TableModel table, IEntity column);

    }
}
