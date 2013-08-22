using System.Web.Mvc;

namespace RCVS.Classes
{
	public class Subject : SelectListItem
	{
		public new string Value { get; set; }
		public new string Text { get; set; }
	}
}