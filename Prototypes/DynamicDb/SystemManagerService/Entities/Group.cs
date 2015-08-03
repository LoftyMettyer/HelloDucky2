using System.ComponentModel.DataAnnotations;

namespace SystemManagerService.Entities
{
    public class Group
    {
        public int Id { get; set; }

        [StringLength(255)]
        public string Name { get; set; }

        public string Description { get; set; }
    }
}
