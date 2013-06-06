using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Fusion.Messages.SocialCare.Schemas.Patterns
{
    public interface IAddressData : IFusionXmlDto
    {
        string workMobile
        {
            get;
            set;
        }

        string personalMobile
        {
            get;
            set;
        }

        string email
        {
            get;
            set;
        }

        string personalEmail
        {
            get;
            set;
        }

        string addressLine1
        {
            get;
            set;
        }

        string addressLine2
        {
            get;
            set;
        }

        string addressLine3
        {
            get;
            set;
        }

        string addressLine4
        {
            get;
            set;
        }

        string addressLine5
        {
            get;
            set;
        }

        string postCode
        {
            get;
            set;
        }

        string homePhoneNumber
        {
            get;
            set;
        }

        string companyName
        {
            get;
            set;
        }
    }
}
