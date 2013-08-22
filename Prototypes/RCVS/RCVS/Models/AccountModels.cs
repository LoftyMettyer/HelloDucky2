using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Globalization;
using System.Web.Mvc;
using System.Web.Security;

namespace RCVS.Models
{
	public class RegisterModel
	{
		[Required]
		[DisplayName("Title")]
		public string Title { get; set; }

		[Required]
		[DisplayName("All first names")]
		public string Forenames { get; set; }

		[Required]
		[DisplayName("All surnames")]
		public string Surnames { get; set; }

		[Required]
		[DataType(DataType.EmailAddress)]
		[DisplayName("Email address")]
		public string EmailAddress { get; set; }

		[DataType(DataType.EmailAddress)]
		[DisplayName("Confirm email address")]
		[Compare("EmailAddress", ErrorMessage = "The email address and confirmation email address do not match.")]
		public string ConfirmEmailAddress { get; set; }

		[Required]
		[StringLength(100, ErrorMessage = "The {0} must be at least {2} characters long.", MinimumLength = 6)]
		[DataType(DataType.Password)]
		public string Password { get; set; }

		[DataType(DataType.Password)]
		[Display(Name = "Confirm password")]
		[Compare("Password", ErrorMessage = "The password and confirmation password do not match.")]
		public string ConfirmPassword { get; set; }

		[Required]
		[DisplayName("Day")]
		public IEnumerable<SelectListItem> Days { get; set; }

		[Required]
		[DisplayName("Month")]
		public IEnumerable<SelectListItem> Months { get; set; }

		[Required]
		[DisplayName("Year")]
		public IEnumerable<SelectListItem> Years { get; set; }

		[Required]
		[DisplayName("Address line 1")]
		public string AddressLine1 { get; set; }

		[Required]
		[DisplayName("Address line 2")]
		public string AddressLine2 { get; set; }

		[DisplayName("Address line 3")]
		public string AddressLine3 { get; set; }

		[Required]
		[DisplayName("Postcode (UK)")]
		public string Postcode { get; set; }

		[Required]
		public string City { get; set; }

		public string County { get; set; }

		[Required]
		[DisplayName("Country")]
		public IEnumerable<SelectListItem> Countries { get; set; }
	}

	public class LoginModel
	{
		[Required]
		[Display(Name = "User name")]
		public string UserName { get; set; }

		[Required]
		[DataType(DataType.Password)]
		[Display(Name = "Password")]
		public string Password { get; set; }
	}

	#region "NOT USED"
	public class UsersContext : DbContext
	{
		public UsersContext()
		//			: base("DefaultConnection")
		{
		}

		public DbSet<UserProfile> UserProfiles { get; set; }
	}

	[Table("UserProfile")]
	public class UserProfile
	{
		[Key]
		[DatabaseGeneratedAttribute(DatabaseGeneratedOption.Identity)]
		public int UserId { get; set; }
		public string UserName { get; set; }
	}
	#endregion
}
