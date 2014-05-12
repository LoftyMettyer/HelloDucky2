using Microsoft.AspNet.Identity.EntityFramework;

namespace ABS_Self_Service.Models
{
    // You can add profile data for the user by adding more properties to your ApplicationUser class, please visit http://go.microsoft.com/fwlink/?LinkID=317594 to learn more.
    public class ApplicationUser : IdentityUser
    {
        public string OpenHRID { get; set; }
        public string OpenPeopleID { get; set; }
        public string OpenAccountsID { get; set; }
    }

    public class ApplicationDbContext : IdentityDbContext<ApplicationUser>
    {
        public ApplicationDbContext()
            : base("DefaultConnection")
        {
        }
    }
}