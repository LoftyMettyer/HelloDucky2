
namespace Fusion.Core.Test
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Xml.Linq;
    using log4net;

    public class QuickOutboundMessageWatcher
    {
        private ILog Logger;

        private string path;

        public QuickOutboundMessageWatcher(string path)
        {
            this.path = path;
            Logger = LogManager.GetLogger(GetType());

            DirectoryUtil.EnforceDirectory(path);

            fsw = new FileSystemWatcher(path, "*.xml");
            fsw.Created += new FileSystemEventHandler(fsw_Created);
            fsw.Renamed += new RenamedEventHandler(fsw_Renamed);
        }

        void fsw_Renamed(object sender, RenamedEventArgs e)
        {
            Logger.InfoFormat("{0}: File renamed from {1} to {2}", this.path, e.OldName, e.Name);
            if (Path.GetExtension(e.Name) == ".xml")
                CopyMessage(e.FullPath);
        }

        void fsw_Created(object sender, FileSystemEventArgs e)
        {
            Logger.InfoFormat("{0}: File created {1}", this.path, e.Name);

            CopyMessage(e.FullPath);
        }

        public IEnumerable<OutboundWatcherDefinition> WatcherDefinitions
        {
            get;
            set;
        }

        private void CopyMessage(string path) {
            if (!File.Exists(path))
                return;

            Logger.InfoFormat("Copying {0} to the relevant out folder...", path);

            try
            {
                FileSharingUtil.WrapSharingViolations(() =>
                {
                    var pathName = FindMessagePath(path);
                    FileUtil.SafeMove(path, Path.Combine(this.path, pathName, Path.GetFileName(path)));
                });
            }
            catch (Exception ex)
            {
                Logger.ErrorFormat("Cannot handle file {0}. {1}", path, ex.Message);
                return;
            }
            
        }

        public void Start()
        {
            // Clear and send existing messages

            Logger.InfoFormat("Scanning path {0} for existing messages", path);

            string[] files = Directory.GetFiles(this.path, "*.xml");

            foreach (string file in files)
            {
                CopyMessage(file);
            }

            Logger.InfoFormat("Now watching {0} for new messages", path);

            fsw.EnableRaisingEvents = true;
        }

        public void Stop()
        {
            Logger.InfoFormat("Stopping watching {0}", path);
            
            fsw.EnableRaisingEvents = false;
        }

        private string FindMessagePath(string file)
        {
            string xml = File.ReadAllText(file);

            var dom = XElement.Parse(xml);

            string targetName = dom.Name.LocalName.Substring(0, 1).ToUpper() + dom.Name.LocalName.Substring(1);

            foreach (OutboundWatcherDefinition wd in this.WatcherDefinitions)
            {
                if (wd.MessageType.Name.StartsWith(targetName))
                {
                    return wd.PathToWatch;
                }
            }
            
            throw new ApplicationException("Message type not found for XML element: " + dom.Name.LocalName + "."); 
        }

        FileSystemWatcher fsw;
    }
}
