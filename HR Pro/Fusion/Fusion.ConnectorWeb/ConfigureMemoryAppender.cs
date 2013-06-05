using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using log4net.Repository.Hierarchy;
using log4net;
using log4net.Appender;
using log4net.Core;

namespace Fusion.ConnectorWeb
{
    public class ConfigureMemoryAppender
    {
        const string AppenderName = "WebLogAppender";

        public void Start()
        {
            //First create and configure the appender  
            MemoryAppender memoryAppender = new MemoryAppender();
            memoryAppender.Name = AppenderName;
            memoryAppender.Threshold = Level.Info;
            //Notify the appender on the configuration changes  
            memoryAppender.ActivateOptions();

            //Get the logger repository hierarchy.  
            log4net.Repository.Hierarchy.Hierarchy repository =
               LogManager.GetRepository() as Hierarchy;

            //and add the appender to the root level  
            //of the logging hierarchy  
            repository.Root.AddAppender(memoryAppender);

            //configure the logging at the root.  
            repository.Root.Level = Level.All;

            //mark repository as configured and  
            //notify that is has changed.  
            repository.Configured = true;
            repository.RaiseConfigurationChanged(EventArgs.Empty);  
        }

        public string[] ReadTest() {
            // Get the default hierarchy
            Hierarchy h = LogManager.GetRepository() as Hierarchy;

            // Get the appender named "MemoryAppender" from the <root> logger
            MemoryAppender ma = h.Root.GetAppender(AppenderName) as
            MemoryAppender;
            
            // Get the events out of the memory appender
            LoggingEvent[] events = ma.GetEvents();

            return events.Select(x => x.RenderedMessage).ToArray() ;

            //return null;
        }
    }
}
