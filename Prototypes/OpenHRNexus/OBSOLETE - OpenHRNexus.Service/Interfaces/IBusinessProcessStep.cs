using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenHRNexus.Service.Interfaces
{
    interface IBusinessProcessStep
    {
        void Submit();
        void PopulateWithData();
        void TranslateTo(string taragetLanguage);
    }
}
