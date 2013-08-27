using System;
using System.ComponentModel;

namespace RCVS.Structures
{
	public struct TRFDetails
	{
		[DisplayName("Date of test")]
		public DateTime DateOfTest { get; set; }

		[DisplayName("Overall band score")]
		public int BandScore { get; set; }
	}
}