using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data.Entity;

namespace SystemManagerService.Entities
{
    public class PermissionGroup
    {
        public int Id { get; set; }

        [StringLength(255)]
        public string Name { get; set; }

        public string Description { get; set; }

        public virtual ICollection<PermissionItem> PermissionItems { get; set; }

    }
}
