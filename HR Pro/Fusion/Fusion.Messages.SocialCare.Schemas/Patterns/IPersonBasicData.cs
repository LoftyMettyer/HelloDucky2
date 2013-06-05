
using System;
namespace Fusion.Messages.SocialCare.Schemas.Patterns
{
    public interface IPersonBasicData : IFusionXmlDto
    {
        string title
        {
            get;
            set;
        }

        string forenames
        {
            get;
            set;
        }

        string surname
        {
            get;
            set;
        }

        string preferredName
        {
            get;
            set;
        }

        bool preferredNameSpecified
        {
            get;
            set;
        }

        gender gender
        {
            get;
            set;
        }

        System.DateTime DOB
        {
            get;
            set;
        }
    }

    public static class PersonBasicDataExtensions
    {
        public static void SetRequiredMembersToEmpty(this IPersonBasicData data)
        {
            data.title = "";
            data.forenames = "";
            data.surname = "";
            data.gender = gender.Unknown;
            data.DOB = DateTime.Now;
        }
    }
}
