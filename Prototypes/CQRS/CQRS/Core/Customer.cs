using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Core
{
	[Table("Customer")]
	public class Customer
	{
		[Key]
		[DatabaseGenerated(DatabaseGeneratedOption.Identity)]
		public int Id { get; set; }

		[Column("Name")]
		public string Name { get; set; }

		[Column("Address")]
		public string Address { get; set; }
	}
}