﻿using System.ComponentModel;
using System.Web;

namespace RCVS.Classes
{
	public class Degree
	{
		[DisplayName("")]
		public string Name { get; set; }

		[DisplayName("Upload your veterinary degree here")]
		public HttpPostedFileBase Document { get; set; }

	}
}