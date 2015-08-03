using System;
using SystemManagerService.Enums;
using SystemManagerService.Interfaces;

namespace SystemManagerService.Messages
{
    public class StructureChangeMessage : IModifyMessage
    {
        public SaveStatusEnum status { get; set; }

    }
}
