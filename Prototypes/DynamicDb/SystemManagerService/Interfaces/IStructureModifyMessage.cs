using SystemManagerService.Enums;

namespace SystemManagerService.Interfaces
{
    public interface IModifyMessage
    {
        SaveStatusEnum status { get; set; }
    }
}
