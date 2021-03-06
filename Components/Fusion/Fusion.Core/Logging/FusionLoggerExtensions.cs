﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.General;

namespace Fusion.Core.Logging
{
    public static class FusionLoggerExtensions
    {
        public static void InfoMessageNonTransactional(this IFusionLogService logService, FusionMessage fm, FusionLogLevel logLevel, string message)
        {
            logService.InfoMessageNonTransactional(fm.Community, fm.Id, logLevel, fm.EntityRef, fm.PrimaryEntityRef, fm.GetMessageName(), message);
        }

        public static void InfoMessageTransactional(this IFusionLogService logService, FusionMessage fm, FusionLogLevel logLevel, string message)
        {
            logService.InfoMessageTransactional(fm.Community, fm.Id, logLevel, fm.EntityRef, fm.PrimaryEntityRef, fm.GetMessageName(), message);
        }

        public static void LogMessageGenerated(this IFusionLogService logService, Messages.General.FusionMessage message)
        {
            logService.InfoMessageNonTransactional(message.Community, message.Id, FusionLogLevel.Info, message.EntityRef, message.PrimaryEntityRef, message.GetMessageName(), "Generating message starting");
            logService.InfoMessageTransactional(message.Community, message.Id, FusionLogLevel.Info, message.EntityRef, message.PrimaryEntityRef, message.GetMessageName(), "Generating message starting (transaction completed)");
        }

        public static void LogMessageReceived(this IFusionLogService logService, Messages.General.FusionMessage message)
        {
            logService.InfoMessageNonTransactional(message.Community, message.Id, FusionLogLevel.Info, message.EntityRef, message.PrimaryEntityRef, message.GetMessageName(), "Message being processed inbound");
            logService.InfoMessageTransactional(message.Community, message.Id, FusionLogLevel.Info, message.EntityRef, message.PrimaryEntityRef, message.GetMessageName(), "Message being processed inbound (transaction completed)");
        }
    }
}
