using System;
using System.Collections.Generic;
using System.Globalization;
using System.Web.Mvc;
using RCVS.Classes;
using RCVS.WebServiceClasses;

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

		public static List<Day> ListOfDays()
		{
			var list = new List<Day>();
			list.Add(new Day { Value = "", Text = "" });


			for (int i = 1; i <= 31; i++)
			{
				list.Add(new Day { Value = i.ToString(), Text = i.ToString() });
			}

			return list;
		}

		public static List<Month> ListOfMonths()
		{
			var list = new List<Month>();
			list.Add(new Month { Value = "", Text = "" });

			for (int i = 0; i <= 11; i++)
			{
				list.Add(new Month
					{
						Value = (i + 1).ToString(),
						Text = CultureInfo.CurrentUICulture.DateTimeFormat.MonthNames[i]
					});
			}

			return list;
		}

		public static List<Year> ListOfYears()
		{
			var list = new List<Year>();
			list.Add(new Year { Value = "", Text = "" });
			for (int i = DateTime.Now.Year; i >= 1900; i--)
			{
				list.Add(new Year { Value = i.ToString(), Text = i.ToString() });
			}

			return list;
		}

		public static void AddActivity(
						long ContactNumber,
						string Activity,
						string ActivityValue,
						string Notes,
						DateTime? ActivityDate,
						string Source
					)
		{
			var client = new IRISWebServices.NDataAccessSoapClient();
			var xmlHelper = new XMLHelper(); //XML helper to serialize and deserialize objects

			var addActivityParameters = new AddActivityParameters
			{
				ContactNumber = ContactNumber,
				Activity = Activity,
				ActivityValue = ActivityValue,
				Notes = Notes,
				Source = Source
			};

			//Dates: I hate them
			if (!ActivityDate.Equals(DateTime.MinValue))
			{
				addActivityParameters.ActivityDate = ActivityDate;
			}

			var serializedParameters = xmlHelper.SerializeToXml(addActivityParameters);
			var response = client.AddActivity(serializedParameters);
			client.Close();
		}

		public static int ActivityIndex(List<SelectContactData_CategoriesResult> ActivityList, string ActivityCode)
		{
			return ActivityList.FindIndex(activity => activity.ActivityCode == ActivityCode);
		}
	}
}