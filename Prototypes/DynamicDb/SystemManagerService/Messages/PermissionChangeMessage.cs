using System;
using SystemManagerService.Enums;
using SystemManagerService.Interfaces;

namespace SystemManagerService.Messages
{
    public class PermissionChangeMessage : IModifyMessage
    {
        public int ModifiedId { get; set; }

        public SaveStatusEnum status { get; set; }
    }
}
