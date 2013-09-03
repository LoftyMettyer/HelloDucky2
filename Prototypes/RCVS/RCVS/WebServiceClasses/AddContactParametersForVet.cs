using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace RCVS.WebServiceClasses
{
	public class AddContactParametersforVet
	{
		public string UserName { get; set; }//Same as email address
		public string Password { get; set; }
		public string EmailAddress { get; set; }
		public string LabelName { get; set; }
		public string Salutation { get; set; }
		public string Title { get; set; }
		public string Surname { get; set; }
		public string Forenames { get; set; }
		public DateTime DateOfBirth { get; set; }
		public string Address { get; set; }
		public string Town { get; set; }
		public string County { get; set; }
		public string Country { get; set; }
		public string Postcode { get; set; }
		public string Source { get; set; }
		public string Status { get; set; }
		public string Position { get; set; }
		public long AddressNumber { get; set; }
		public long OrganisationNumber { get; set; }
	}
}