
namespace Fusion.Core.Test
{
    using System;
    using System.IO;
    using Fusion.Core.MessageSenders;
    using Fusion.Messages.General;
    using log4net;
    using StructureMap.Attributes;

    public class OutboundMessageWatcher<T> : IOutboundMessageWatcher where T : FusionMessage, new()
    {
        private ILog Logger;

        [SetterProperty]
        public IMessageSenderInvoker MessageSenderInvoker
        {
            get;
            set;
        }

        [SetterProperty]
        public IFusionXmlMetadataExtractInvoker FusionXmlMetadataExtractInvoker
        {
            get;
            set;
        } 


        private string path;
        private string messageType;
        private string communityName;
        private IFusionXmlMetadataExtract metadataExtractor;


        /// <summary>
        /// Gets or sets the message sender (rather than searching)
        /// </summary>
        /// <value>
        /// The message sender.
        /// </value>
        public IMessageSender MessageSender
        {
            get;
            set;
        }

        public OutboundMessageWatcher(IFusionXmlMetadataExtract metadataExtractor, string path, string communityName)
        {
            this.path = path;
            this.messageType = this.GetType().GetGenericArguments()[0].Name;
            this.communityName = communityName;
            this.metadataExtractor = metadataExtractor;

            Logger = LogManager.GetLogger(GetType());

            DirectoryUtil.EnforceDirectory(this.path);

            fsw = new FileSystemWatcher(this.path, "*.xml");
            fsw.Created += new FileSystemEventHandler(fsw_Created);
            fsw.Renamed += new RenamedEventHandler(fsw_Renamed);
        }

        void fsw_Renamed(object sender, RenamedEventArgs e)
        {
            Logger.InfoFormat("{0}: File renamed from {1} to {2}", this.path, e.OldName, e.Name);
            if (Path.GetExtension(e.Name) == ".xml")
                SendMessage(e.FullPath);
        }

        void fsw_Created(object sender, FileSystemEventArgs e)
        {
            Logger.InfoFormat("{0}: File created {1}", this.path, e.Name);

            SendMessage(e.FullPath);
        }


        private void SendMessage(string path) {
            if (!File.Exists(path))
                return;

            Logger.InfoFormat("Sending {0}:{1}", messageType, path);

            T msg = new T
            {
                Originator = "Test",                
                Id = Guid.NewGuid(),
                Community = this.communityName,
                CreatedUtc = DateTime.UtcNow
            };

            try
            {
                FileSharingUtil.WrapSharingViolations(() =>
                {
                    msg.Xml = File.ReadAllText(path);
                });
            }
            catch
            {
                Logger.ErrorFormat("Cannot read file {0}", path);
                return;
            }

            var xmlMetaData = this.metadataExtractor.GetMetadataFromXml(msg);

            msg.SchemaVersion = xmlMetaData.Version;
            msg.EntityRef = xmlMetaData.EntityRef;
            msg.PrimaryEntityRef = xmlMetaData.PrimaryEntityRef;

            if (this.MessageSender != null)
            {
                this.MessageSender.Send(msg);
            }
            else
            {
                MessageSenderInvoker.Invoke(msg);
            }

            FileUtil.SafeMove(path, Path.ChangeExtension(path, ".sent"));
        }

        FileSystemWatcher fsw;

        public void Start()
        {
            // Clear and send existing messages

            Logger.InfoFormat("Scanning path {0} for existing messages of type {1}", path, messageType);

            string[] files = Directory.GetFiles(this.path, "*.xml");

            foreach (string file in files)
            {
                SendMessage(file);
            }

            Logger.InfoFormat("Now watching {0} for new messages of type {1}", path, messageType);

            fsw.EnableRaisingEvents = true;
        }

        public void Stop()
        {
            Logger.InfoFormat("Stopping watching {0}", path);
            
            fsw.EnableRaisingEvents = false;
        }


    }
}
