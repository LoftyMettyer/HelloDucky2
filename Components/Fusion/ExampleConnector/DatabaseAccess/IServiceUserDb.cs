using System;
namespace Connector1.DatabaseAccess
{
    public interface IServiceUserDb
    {

        string MessageContext
        {
            get;
            set;
        }
        int CreateServiceUser(string forename, string surname);
        ServiceUser ReadServiceUser(int id);
        void UpdateServiceUser(int id, string forename, string surname);
    }
}
