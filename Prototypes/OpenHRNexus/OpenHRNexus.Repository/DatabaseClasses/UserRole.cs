﻿using System;
using System.ComponentModel.DataAnnotations;

namespace OpenHRNexus.Repository.DatabaseClasses
{
	public class UserRole
	{
		[Key]
		public int Id { get; set; }
		public Guid UserID { get; set; }
		public string Name { get; set; }
	}
}
