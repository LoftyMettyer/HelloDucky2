using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NServiceBus.Hosting.Profiles;
using NServiceBus;
using NServiceBus.Hosting;
using StructureMap.Attributes;

namespace Fusion.Publisher.SocialCare.Staff
{
    //public class ScriptDatabase : IProfile
    //{
    //}

    //public class ScriptDatabaseHandler : IHandleProfile<ScriptDatabase>
    //{
    //    public void ProfileActivated()
    //    {
    //        //NServiceBus.Host.Internal.GenericHost.ConfigurationComplete += (o, e) =>
    //        //    {

    //        var obj = Configure.Instance.Builder.Build<INHibernateSession>();
    //        obj.ScriptDatabase();
    //        //}

    //    }
    //}
    // 
    public class ScriptDatabase : IWantToRunAtStartup
    {

        [SetterProperty]
        public INHibernateSession Session
        {
            get;
            set;
        }

        public void Run()
        {
            //this.Session.ScriptDatabase();

        }

        public void Stop()
        {
            
        }
    }
}
