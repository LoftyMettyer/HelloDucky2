using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DynamicDb
{
    public class SystemManager
    {
        public void CreateTables()
        {
            var blah = new DynamicODataService();
            blah.CreateTable("SomeTable");
        }

    }
}
