using System;

namespace DayPilot.Utils
{
	public static class Extensions
	{

		public static bool IsWeekend(this DateTime instance) {
			var day = instance.DayOfWeek;
			return !((day >= DayOfWeek.Monday) && (day <= DayOfWeek.Friday));
		}

	}
}
