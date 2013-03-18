namespace Fusion.Core.Test
{
    using System;
    using System.IO;
    using Fusion.Core;
    using Fusion.Messages.General;
    using log4net;
    using StructureMap.Attributes;

    public class BaseWriteFileMessageHandler
    {
        protected ILog Logger
        {
            get;
            set;
        }

        [SetterProperty]
        public ITestingConfiguration TestingConfiguration
        {
            get;
            set;
        }

        public BaseWriteFileMessageHandler()
        {
            Logger = LogManager.GetLogger(this.GetType());
        }

        public void WriteMessage(FusionMessage message)
        {
            string name = message.GetType().Name;

            Logger.Info(string.Format("Test connector received " + name + " with Id {0} from {1}.", message.Id, message.Originator));

            DirectoryUtil.EnforceDirectory(Path.Combine(TestingConfiguration.MessagePath, "in", name));

            string path = Path.Combine(TestingConfiguration.MessagePath, "in", name, String.Format("{0} {1} {2}.xml", 
                message.GetMessageName(), 
                message.EntityRef.HasValue ? message.EntityRef.Value.ToString() : "-", message.Id));
            
            File.WriteAllText(path, message.Xml);
        }

    }
}
