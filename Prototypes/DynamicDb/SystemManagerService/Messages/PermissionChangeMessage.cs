using SystemManagerService.Enums;
using SystemManagerService.Interfaces;

namespace SystemManagerService.Messages
{
    public class PermissionChangeMessage : IModifyMessage
    {
        public SaveStatusEnum status { get; set; }
    }
}
