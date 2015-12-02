using System.Data.Entity;

namespace OpenHRTestToLive
{
    public class MyDbContext : DbContext
    {
        public MyDbContext(string nameOrConnectionString)
            : base(nameOrConnectionString)
        {
        }
    }
}
