using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace SystemManagerService.Entities
{
    public class PermissionCategory
    {
        public int Id { get; set; }

        [StringLength(25)]
        public string KeyName { get; set; }

        [StringLength(255)]
        public string Description { get; set; }

        public virtual IEnumerable<PermissionItem> Items { get; set; }


    }
}
