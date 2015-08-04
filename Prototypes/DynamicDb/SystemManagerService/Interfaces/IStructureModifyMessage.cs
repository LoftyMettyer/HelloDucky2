using SystemManagerService.Enums;

namespace SystemManagerService.Interfaces
{
    public interface IModifyMessage
    {
        int ModifiedId { get; set; }
        SaveStatusEnum status { get; set; }

        //string Detail { get; set; }
    }
}
