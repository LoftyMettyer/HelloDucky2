using System.Data.Entity;

namespace OpenHR.TestToLive
{
    public class MyDbContext : DbContext
    {
        public MyDbContext(string nameOrConnectionString)
            : base(nameOrConnectionString)
        {
        }
    }
}
