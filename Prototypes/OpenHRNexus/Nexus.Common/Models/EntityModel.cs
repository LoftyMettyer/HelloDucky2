using System.ComponentModel.DataAnnotations;

namespace Nexus.Common.Models
{
    public class EntityModel
    {

        public int Id { get; set; }

        [MaxLength(250)]
        public string Name { get; set; }
    }
}
