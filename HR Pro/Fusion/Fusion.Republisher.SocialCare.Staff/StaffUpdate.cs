using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.SocialCare;
using Fusion.Publisher.SocialCare.Database;

namespace Fusion.Publisher.SocialCare
{
    public class StaffUpdate
    {

        public StaffUpdate(IStaffRepository staffRepository)
	    {
            this.staffRepository = staffRepository;
	    }

        IStaffRepository staffRepository;

        public void UpdateFrom(StaffChangeMessage message)
        {
           

        }
    }
}
