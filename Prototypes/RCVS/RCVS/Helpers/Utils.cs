using System;
using System.Collections.Generic;
using System.Globalization;
using System.Web.Mvc;

namespace RCVS.Helpers
{
	public static class Utils
	{

		public enum DropdownListType
		{
			Days,
			Months,
			Years
		}
		public static List<SelectListItem> DropdownList(DropdownListType type)
		{
			int i;
			var list = new List<SelectListItem>();
			list.Add(new SelectListItem { Value = "", Text = "" });

			if (type == DropdownListType.Days)
				for (i = 1; i <= 31; i++)
				{
					list.Add(new SelectListItem { Value = i.ToString(), Text = i.ToString() });
				}
			else if (type == DropdownListType.Months)
			{
				for (i = 0; i <= 11; i++)
				{
					list.Add(new SelectListItem { Value = (i + 1).ToString(), Text = CultureInfo.CurrentUICulture.DateTimeFormat.MonthNames[i] });
				}
			}
			else
			{
				for (i = DateTime.Now.Year; i >= 1900; i--)
				{
					list.Add(new SelectListItem { Value = i.ToString(), Text = i.ToString() });
				}
			}

			return list;
		}
	}
}