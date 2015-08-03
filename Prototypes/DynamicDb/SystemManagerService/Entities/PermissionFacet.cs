using System.ComponentModel.DataAnnotations;

namespace SystemManagerService.Entities
{
    public class PermissionFacet
    {
        public int Id { get; set; }

        [StringLength(25)]
        public string Name { get; set; }
    }
}
