using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SystemFramework.ErrorHandler;
using SystemFramework.Forms;

namespace System_Framework.Test
{
	class Program
	{
		[STAThread]
		static void Main(string[] args)
		{
			SystemFramework.Globals.Initialise();

			var errors = SystemFramework.Globals.ErrorLog;
			errors.Add(Section.TableAndColumns, "ObjectName", Severity.Warning, "Error creating calculation for Second_Holiday_Requests.Holiday_Authorised_EMail ", "Detail");
			errors.Add(Section.TableAndColumns, "ObjectName", Severity.Warning, "Message", "Detail");
			errors.Add(Section.TableAndColumns, "ObjectName2", Severity.Warning, "Message", "Detail2");
			errors.Add(Section.TableAndColumns, "ObjectName2", Severity.Warning, "Message", "Detail");
			errors.Add(Section.TableAndColumns, "ObjectName3", Severity.Warning, "Message", "Detail");
			errors.Add(Section.TableAndColumns, "ObjectName3", Severity.Warning, "Message", "Detail999");

			using (var f = new ErrorLog()) {
				f.ShowDialog();
			}
		}
	}
}
