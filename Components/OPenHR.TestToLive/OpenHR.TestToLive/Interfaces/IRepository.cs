using OpenHR.TestToLive.Enums;

namespace OpenHR.TestToLive.Interfaces
{
    public interface IRepository
    {
        void Connection(string connection);
        string ExportDefinition(int Id);
        RepositoryStatus ImportDefinitions();
    }
}
