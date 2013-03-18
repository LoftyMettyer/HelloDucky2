
using NServiceBus.Installation;
using System.Configuration;
using NServiceBus.Utils;
using System;
using NServiceBus.Installation.Environments;
using System.Messaging;
using NServiceBus;

class InstallAuditQueue : INeedToInstallSomething<Windows>
{
    public void Install(System.Security.Principal.WindowsIdentity identity)
    {
        var auditMessagesQueueName = ConfigurationManager.AppSettings["AuditQueue"];
        Console.WriteLine("Checking to see if Audit Queue: {0} exists ...", auditMessagesQueueName);
        var auditMessagesQueuePath = MsmqUtilities.GetFullPathWithoutPrefix(auditMessagesQueueName);
        if (!MessageQueue.Exists(auditMessagesQueuePath))
        {
            Console.WriteLine("Creating Audit Queue: {0}", ConfigurationManager.AppSettings["AuditQueue"]);
            MessageQueue.Create(auditMessagesQueuePath, true);
        }

        //NServiceBus.Utils.MsmqUtilities.CreateQueueIfNecessary(Address.Local.SubScope("auditqueue"), identity.User);           

    }
}