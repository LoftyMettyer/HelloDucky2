using SystemManagerService.Interfaces;
using SystemManagerService.Messages;

namespace SystemManagerService
{
    public class Structure
    {
        public IModifyMessage AddTable(string name)
        {
            var result = new StructureChangeMessage();
            return result;
        }

        public IModifyMessage AddColumn(string name, int type)
        {
            var result = new StructureChangeMessage();
            return result;
        }

    }
}
