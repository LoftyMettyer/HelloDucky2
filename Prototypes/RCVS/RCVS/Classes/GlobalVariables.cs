using System;
using System.Web.Hosting;

namespace RCVS.Classes
{
	public static class GlobalVariables
	{
		public static string LogFileFullPath
		{
			get { return HostingEnvironment.MapPath("/Logs/WebServiceCalls_" + DateTime.Now.Date.ToString().Substring(0,10).Replace("/", "-") + ".txt"); }
		}
	}
}