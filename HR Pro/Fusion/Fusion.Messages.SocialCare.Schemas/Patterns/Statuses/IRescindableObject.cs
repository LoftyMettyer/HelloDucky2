namespace Fusion.Messages.SocialCare.Schemas.Patterns.Statuses
{
    public interface IRescindableRecord
    {
        recordStatusRescindable recordStatus
        {
            get;
            set;
        }

        bool recordStatusSpecified
        {
            get;
            set;
        }
    }
}
