using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Fusion.Core.Sql
{
    public class BusTranslationResults
    {
        public Guid? BusRef
        {
            get;
            set;
        }

        public bool HasValue
        {
            get
            {
                return BusRef.HasValue;
            }
        }

        public Guid Value
        {
            get
            {
                return BusRef.Value;
            }
        }

        public string LocalId
        {
            get;
            set;
        }

        public bool BusRefNewlyCreated
        {
            get;
            set;
        }
    }
}
