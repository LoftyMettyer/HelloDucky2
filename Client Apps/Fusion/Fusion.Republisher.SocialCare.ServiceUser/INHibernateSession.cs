using System;
namespace Fusion.Publisher.SocialCare.Staff
{
    public interface INHibernateSession
    {
        NHibernate.ISession OpenSession();
        void ScriptDatabase();
    }
}
