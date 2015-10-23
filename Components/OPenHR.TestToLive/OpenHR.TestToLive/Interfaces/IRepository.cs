using OpenHR.TestToLive.Enums;

namespace OpenHR.TestToLive.Interfaces
{
    public interface IRepository
    {
        void Connection(string username, string password, string database, string server);
        string ExportDefinition(int Id);
        RepositoryStatus ImportDefinitions();
    }
}
